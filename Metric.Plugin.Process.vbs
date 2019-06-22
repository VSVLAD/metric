Option Explicit

' Обязательная функция для иницилазации класса
Public Function NewClass
    Set NewClass = New MetricProcess
End Function

Class MetricProcess

    ' Название метрики
    Public Property Get Name
        Name = "process"
    End Property

    ' Пересоздание схемы
    Public Function RecreateTable(PgString)
        ModifyPgData " drop table if exists metric_process;                             " & _
                     " create table metric_process(id bigserial,                        " & _
                     "                             host_name varchar(50) not null,      " & _
                     "                             metric_name varchar(50) not null,    " & _
                     "                             metric_date timestamp default now(), " & _
                     "                             Name varchar(200),                   " & _
                     "                             ProcessId integer,                   " & _
                     "                             Priority integer,                    " & _
                     "                             ExecutablePath varchar(1024),        " & _
                     "                             CommandLine varchar(2048),           " & _
                     "                             HandleCount integer,                 " & _
                     "                             ThreadCount integer,                 " & _
                     "                             UserModeTime bigint,                 " & _
                     "                             VirtualSize bigint,                  " & _
                     "                             WorkingSetSize bigint,               " & _
                     "                             ReadOperationCount bigint,           " & _
                     "                             WriteOperationCount bigint)          ", PgString
                     '
    End Function

    ' Добавление свежих метрик
    Public Function InsertRows(HostName, Params, PgString, WmiString)
        Dim item, colItems, timestampNow

        colItems = SelectWmiData(" select Name, ProcessId, Priority, ExecutablePath, CommandLine, HandleCount, ThreadCount, UserModeTime, VirtualSize, WorkingSetSize, ReadOperationCount, WriteOperationCount from Win32_Process ", WmiString)
        timestampNow = toDate(Now)

        For Each item In colItems
            ModifyPgData " insert into metric_process(host_name, metric_name, metric_date, Name, ProcessId, Priority, ExecutablePath, CommandLine, HandleCount, ThreadCount, UserModeTime, VirtualSize, WorkingSetSize, ReadOperationCount, WriteOperationCount) values (" & _
                                toStr(HostName) 		                                                                          & " ," & _
                                toStr("process") 		                                                                          & " ," & _
                                ToStr(timestampNow)		                                                                          & " ," & _
                                toStr(item("Name"))		                                                                          & " ," & _
                                ToReal(item("ProcessId")) 		                                                                  & " ," & _
                                ToReal(item("Priority")) 			                                                              & " ," & _
                                toStr(item("ExecutablePath"))                                                                     & " ," & _
                                toStr(item("CommandLine"))                                                                        & " ," & _
                                ToReal(item("HandleCount"))                                                                       & " ," & _
                                ToReal(item("ThreadCount"))                                                                       & " ," & _
                                ToReal(item("UserModeTime"))                                                                      & " ," & _
                                ToReal(item("VirtualSize"))                                                                       & " ," & _
                                ToReal(item("WorkingSetSize"))                                                                    & " ," & _
                                ToReal(item("ReadOperationCount"))                                                                & " ," & _
                                ToReal(item("WriteOperationCount"))                                                               & " )", PgString
        Next
    End Function

End Class