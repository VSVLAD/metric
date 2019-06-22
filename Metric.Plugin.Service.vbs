Option Explicit

' Обязательная функция для иницилазации класса
Public Function NewClass
    Set NewClass = New MetricService
End Function

Class MetricService

    ' Название метрики
    Public Property Get Name
        Name = "service"
    End Property

    ' Пересоздание схемы
    Public Function RecreateTable(PgString)
        ModifyPgData " drop table if exists metric_service;                                           " & _
                     " create table metric_service(id bigserial,                                      " & _
                     "                             host_name varchar(50) not null,                    " & _
                     "                             metric_name varchar(50) not null,                  " & _
                     "                             metric_date timestamp default now(),               " & _
                     "                             Name varchar(200),                                 " & _
                     "                             DisplayName varchar(200),                          " & _
                     "                             PathName varchar(1024),                            " & _
                     "                             ProcessId integer,                                 " & _
                     "                             StartMode varchar(50),                             " & _
                     "                             StartName varchar(50),                             " & _
                     "                             State varchar(50),                                 " & _
                     "                             Status varchar(50),                                " & _
                     "                             ExitCode integer);                                 " & _
                     " create index idx_mservice_date_host on metric_service(host_name, metric_date); ", PgString
    End Function

    ' Добавление свежих метрик
    Public Function InsertRows(HostName, Params, PgString, WmiString)
        Dim item, colItems, timestampNow

        colItems = SelectWmiData(" select Name, DisplayName, PathName, ProcessId, StartMode, StartName, State, Status, ExitCode from Win32_Service ", WmiString)
        timestampNow = toDate(Now)

        For Each item In colItems
            ModifyPgData " insert into metric_service(host_name, metric_name, metric_date, Name, DisplayName, PathName, ProcessId, StartMode, StartName, State, Status, ExitCode) values (" & _
                                toStr(HostName) 		                                                                                                                & " ," & _
                                toStr("service") 		                                                                                                                & " ," & _
                                ToStr(timestampNow)		                                                                                                                & " ," & _
                                toStr(item("Name"))		                                                                                                                & " ," & _
                                toStr(item("DisplayName")) 		                                                                                                        & " ," & _
                                toStr(item("PathName")) 			                                                                                                    & " ," & _
                                ToReal(item("ProcessId"))                                                                                                               & " ," & _
                                toStr(item("StartMode"))                                                                                                                & " ," & _
                                toStr(item("StartName"))                                                                                                                & " ," & _
                                toStr(item("State"))                                                                                                                    & " ," & _
                                toStr(item("Status"))                                                                                                                   & " ," & _
                                ToReal(item("ExitCode"))                                                                                                                & " )", PgString
        Next
    End Function

End Class