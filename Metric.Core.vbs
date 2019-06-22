Option Explicit


' Обязательная функция для иницилазации класса
Public Function NewClass
    Set NewClass = Nothing
End Function

' Пересоздаём схему данных для правил метрик
Public Sub ReinstallMetricRule(PgString)
	ModifyPgData " drop table if exists metric_rule;                                             				     " & _
                " create table metric_rule (	                                                     				     " & _
		"    host_name   varchar(50) not null,	                                             				     " & _
		"    metric_name varchar(50) not null,	                                             				     " & _
                "    params      text null,	                                                     				     " & _
		"    life_time   integer default 86400,		                                     				     " & _
		"    period      integer default 60			                             				     " & _
		" );							                             				     " & _
		" create unique index idx_mrule_host_metric on metric_rule(host_name, metric_name);  				     " & _
                " insert into metric_rule values('localhost', 'cpu', null, '86400', '5'), 	         			     " & _
		"								 ('localhost', 'disk', null, '86400', '30'),	     " & _
                "								 ('localhost', 'service', null, '86400', '30'),      " & _
                "								 ('localhost', 'process', null, '86400', '15');      ", PgString
End Sub

' Возвращает текущую дату с милисекундами, уникальная на каждый вызов
Public Function toDateUniq
    Dim Milliseconds, Seconds, Minutes, Hours
    Dim tmr, temp, strTime

    tmr = Timer
	temp = Int(tmr)
	
    Milliseconds = Int((tmr - temp) * 1000)
    Seconds = temp mod 60
    temp    = Int(temp / 60)
    Minutes = temp Mod 60
    Hours   = Int(temp / 60)

    strTime =           Right("0"    & Hours, 2)        & ":"
    strTime = strTime & Right("0"    & Minutes, 2)      & ":"
    strTime = strTime & Right("0"    & Seconds, 2)      & "."
    strTime = strTime & Right("0000" & Milliseconds, 4)
	
	' Немного ожидания для уникальности
    Do While Timer = tmr
    Loop
	
    toDateUniq = Year(Now) 	  & "-" & _
           Right("0" & Month(Now), 2) & "-" & _
           Right("0" & Day(Now), 2)   & " " & strTime
End Function

' Возвращает дату по формату YYYY-MM-DD HH:mm:ss
Public Function toDate(Value)
    If Not IsNull(Value) Then
        toDate = Year(Value) 	          & "-" & _
            Right("0" & Month(Value), 2)  & "-" & _
            Right("0" & Day(Value), 2)    & " " & _
            Right("0" & Hour(Value), 2)   & ":" & _
            Right("0" & Minute(Value), 2) & ":" & _
            Right("0" & Second(Value), 2)
    Else
        toDate = "null"
    End If
End Function

' Возвращает строку для вставки в запросы SQL
Public Function toStr(Value)
    If Not IsNull(Value) Then
        toStr = "'" & Replace(Value, "'", "''") & "'"
    Else
        toStr = "null"
    End If
End Function

' Возвращает число для вставки в запросы SQL
Public Function toReal(Value)
    If Not IsNull(Value) Then
        toReal = Replace(Value, ",", ".")
    Else
        toReal = "null"
    End If
End Function

' Читаем из Postgres
Public Function SelectPgData(Query, ConnectString)
    On Error Resume Next

    Dim psqlConn
    Set psqlConn = CreateObject("ADODB.Connection")
    psqlConn.Open ConnectString
    
    Dim psqlReader
    Set psqlReader = psqlConn.Execute(Query)
    
    If psqlReader.BOF Or psqlReader.EOF Then
        SelectPgData = Array()
    Else
        SelectPgData = psqlReader.GetRows
    End If
    
    psqlConn.Close

    If Err.Number Then LogMessage "::Metric.SelectPgData [" & Query & ", " & ConnectString & "]"
    On Error GoTo 0
End Function

'Добавляем, изменяем, удаляем в Postgres
Public Function ModifyPgData(Query, ConnectString)
    On Error Resume Next

    Dim psqlConn
    Set psqlConn = CreateObject("ADODB.Connection")
    
    psqlConn.Open ConnectString
    psqlConn.Execute Query
    psqlConn.Close

    If Err.Number Then LogMessage "::Metric.ModifyPgData [" & Query & ", " & ConnectString & "]"
    On Error GoTo 0
End Function

' Открываем соединение Postgres для ручного управления транзакцией
Public Function OpenPgConnection(ConnectString)
    On Error Resume Next

    Dim psqlConn
    Set psqlConn = CreateObject("ADODB.Connection")

    psqlConn.Open ConnectString
    Set OpenPgConnection = psqlConn

    If Err.Number Then LogMessage "::Metric.OpenPgConnection [" & ConnectString & "]"
    On Error GoTo 0
End Function

' Открываем соединение Postgres для ручного управления транзакцией
Public Function ExecuteOnPgConnection(Query, ConnectionObject)
    On Error Resume Next
    
    ConnectionObject.Execute Query

    If Err.Number Then LogMessage "::Metric.ExecuteOnPgConnection [" & Query & "]"
    On Error GoTo 0
End Function

' Читаем из WMI
Public Function SelectWmiData(Query, ConnectString)
    On Error Resume Next

    Dim objWMIService, colItems, objItem, objProp
    Set objWMIService = GetObject(ConnectString)
    Set colItems = objWMIService.ExecQuery(Query, , 48)
    
    ' Массив содержит элементы с типом "Словарь"
    Dim resultArr
    resultArr = Array()
    
    For Each objItem In colItems
    
        'Словарь для текущей итерации
        Dim resultDict
        Set resultDict = CreateObject("Scripting.Dictionary")
        
        ' Увеличиваем массив на единицу, сохраняя предыдущие ссылки
        ReDim Preserve resultArr(UBound(resultArr) + 1)
        Set resultArr(UBound(resultArr)) = resultDict
        
        For Each objProp In objItem.Properties_
            resultDict.Add objProp.Name, objProp.Value
        Next
    Next
    
    SelectWmiData = resultArr

    If Err.Number Then LogMessage "::Metric.SelectWmiData [" & Query & ", " & ConnectString & "]"
    On Error GoTo 0
End Function

' Возвращает коллекцию параметров из коммандной строки [task, host, metric]
Public Function SelectTaskParameters
    On Error Resume Next
    Dim resultDict, argNames

    Set resultDict = CreateObject("Scripting.Dictionary")
    Set argNames = WScript.Arguments.Named

    If argNames.Exists("task")   Then resultDict.Add "task",   argNames("task")   Else resultDict.Add "task",   ""
    If argNames.Exists("host")   Then resultDict.Add "host",   argNames("host")   Else resultDict.Add "host",   "%"
    If argNames.Exists("metric") Then resultDict.Add "metric", argNames("metric") Else resultDict.Add "metric", ""

    Set SelectTaskParameters = resultDict

    If Err.Number Then LogMessage "::Metric.SelectTaskParameters [" & argNames.Count & ", " & resultDict.Count & "]"
    On Error GoTo 0
End Function
