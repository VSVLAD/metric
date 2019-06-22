Option Explicit


' Обязательная функция для иницилазации класса
Public Function NewClass
    Set NewClass = New MetricNginxLog
End Function

Class MetricNginxLog

    ' Название метрики
    Public Property Get Name
        Name = "nginxlog"
    End Property

    ' Пересоздание схемы
    Public Function RecreateTable(PgString)
        ModifyPgData " drop table if exists metric_nginxlog;                            " & _
                     " create table metric_nginxlog(id bigserial,                       " & _
                     "                             host_name varchar(50) not null,      " & _
                     "                             metric_name varchar(50) not null,    " & _
                     "                             metric_date timestamp default now(), " & _
                     "                             FilePath varchar(1024),              " & _
                     "                             RequestIP varchar(50),               " & _
                     "                             RequestDate timestamp,               " & _
                     "                             RequestMethod varchar(10),           " & _
                     "                             RequestUrl varchar(2048),            " & _
                     "                             ResponseHttp varchar(50),            " & _
                     "                             ResponseCode integer,                " & _
                     "                             ResponseLength integer,              " & _
                     "                             UserAgent varchar(500))              ", PgString
    End Function

    ' Добавление свежих метрик
    Public Function InsertRows(HostName, Params, PgString, WmiString)
        Dim timestampNow
        timestampNow = toDate(Now)

        ' По всем файлам из параметра
        Dim fileLog
        For Each fileLog In Split(Params, "|")

            Dim fso, tsFile
            Set fso = CreateObject("Scripting.FileSystemObject")
            If fso.FileExists(fileLog) Then

                LogMessage "::Metric.Plugin.NginxLog fileLog=[" & fileLog & "] выполняется обработка..."
                Set tsFile = fso.OpenTextFile(fileLog, 1, False, 0)  ' 0 - ascii, 1 - unicode

                ' Создаём объект для регулярного выражения
                Dim regExp, regMatches, regSubMatches, lineDate
                Set regExp = CreateObject("VBScript.RegExp")
                regExp.Pattern = "(\S+) (\S+) (\S+) \[([\w:/]+\s[+\-]\d{4})\] ""(\S+)\s?(\S+)?\s?(\S+)?"" (\d{3}|-) (\d+|-)\s?""?([^""]*)""?\s?""?([^""]*)?""?"
                regExp.Global = True
                regExp.IgnoreCase = True
                
                ' Находим последние изменения в базе
                Dim lastRequestDate
                lastRequestDate = SelectPgData("select coalesce(max(RequestDate), '2000-01-01') as max_date from metric_nginxlog where FilePath = " & ToStr(fileLog), PgString)(0, 0)
                lastRequestDate = CDate(lastRequestDate)

                ' Открываем траназакцию
                Dim transactConnect
                Set transactConnect = OpenPgConnection(PgString)
                ExecuteOnPgConnection "begin transaction", transactConnect

                ' По всем строкам проходимся регуляркой
                On Error Resume Next
                Do While Not tsFile.AtEndOfStream
                    Set regMatches = regExp.Execute(tsFile.ReadLine)
                    Set regSubMatches = regMatches(0).SubMatches
                    
                    'Находим последние изменения в логе
                    lineDate = CDate(ToDateConvert(regSubMatches(3)))
                    If lineDate > lastRequestDate Then

                        ExecuteOnPgConnection " insert into metric_nginxlog(host_name, metric_name, metric_date, FilePath, RequestIP, RequestDate, RequestMethod, RequestUrl, ResponseHttp, ResponseCode, ResponseLength, UserAgent) values (" & _
                                                    toStr(HostName) 		                                                        & " ," & _
                                                    toStr("nginxlog") 		                                                        & " ," & _
                                                    toStr(timestampNow)		                                                        & " ," & _
                                                    toStr(fileLog)		                                                            & " ," & _
                                                    ToStr(regSubMatches(0)) 		                                                & " ," & _
                                                    ToStr(lineDate) 			                                                    & " ," & _
                                                    toStr(regSubMatches(4))                                                         & " ," & _
                                                    toStr(regSubMatches(5))                                                         & " ," & _
                                                    toStr(regSubMatches(6))                                                         & " ," & _
                                                    toReal(regSubMatches(7))                                                        & " ," & _
                                                    toReal(regSubMatches(8))                                                        & " ," & _
                                                    toStr(regSubMatches(10))                                                        & " )", transactConnect
                    End If

                    If Err.Number <> 0 Then
                        LogMessage "::Metric.Plugin.NginxLog InsertRows [...]"
                        Err.Clear
                    End If
                Loop
                On Error Goto 0

                ExecuteOnPgConnection "commit", transactConnect
                transactConnect.Close
            Else
                LogMessage "::Metric.Plugin.NginxLog fileLog=[" & fileLog & "] не найден! Пропуск файла"
            End If
        Next
    End Function


    ' Конвертирует формат даты из 01/Jan/2000:12:00:00 в 2000-01-01 12:00:00
    Private Function ToDateConvert(Value)
        Dim arr, arr2

        arr = Split(Value, "/")
        arr2 = Split(arr(2), ":")

        Select Case arr(1)
            Case "Jan": arr(1) = "01"
            Case "Feb": arr(1) = "02"
            Case "Mar": arr(1) = "03"
            Case "Apr": arr(1) = "04"
            Case "May": arr(1) = "05"
            Case "Jun": arr(1) = "06"
            Case "Jul": arr(1) = "07"
            Case "Aug": arr(1) = "08"
            Case "Sep": arr(1) = "09"
            Case "Oct": arr(1) = "10"
            Case "Nov": arr(1) = "11"
            Case "Dec": arr(1) = "12"
        End Select

        ToDateConvert = arr2(0) & "-" & arr(1) & "-" & arr(0) & " " & arr2(1) & ":" & arr2(2) & ":" & Left(arr2(3), 2)
    End Function

End Class
