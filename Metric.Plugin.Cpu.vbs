Option Explicit

' Обязательная функция для иницилазации класса
Public Function NewClass
    Set NewClass = New MetricCpu
End Function

Class MetricCpu

    ' Название метрики
    Public Property Get Name
        Name = "cpu"
    End Property

    ' Пересоздание схемы
    Public Function RecreateTable(PgString)
        ModifyPgData " drop table if exists metric_cpu;                             " & _
                     " create table metric_cpu(id bigserial,                        " & _
                     "                         host_name varchar(50) not null,      " & _
                     "                         metric_name varchar(50) not null,    " & _
                     "                         metric_date timestamp default now(), " & _
                     "                         Name varchar(100),                   " & _
                     "                         NumberOfCores integer,               " & _
                     "                         NumberOfLogicalProcessors integer,   " & _
                     "                         MaxClockSpeed integer,               " & _
                     "                         CurrentClockSpeed integer,           " & _
                     "                         LoadPercentage integer)              ", PgString
    End Function

    ' Добавление свежих метрик
    Public Function InsertRows(HostName, Params, PgString, WmiString)
        Dim item, colItems, timestampNow
        
        colItems = SelectWmiData(" select Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, CurrentClockSpeed, LoadPercentage from Win32_Processor ", WmiString)
        timestampNow = toDate(Now)

        For Each item In colItems
            ModifyPgData " insert into metric_cpu(host_name, metric_name, metric_date, Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, CurrentClockSpeed, LoadPercentage) values (" & _
                                toStr(HostName) 		                                     & " ," & _
                                toStr("cpu") 		                                         & " ," & _
                                ToStr(timestampNow)		                                     & " ," & _
                                ToStr(item("Name"))		                                     & " ," & _
                                ToReal(item("NumberOfCores")) 		                         & " ," & _
                                ToReal(item("NumberOfLogicalProcessors")) 			         & " ," & _
                                ToReal(item("MaxClockSpeed"))                                & " ," & _
                                ToReal(item("CurrentClockSpeed"))                            & " ," & _
                                ToReal(item("LoadPercentage"))                               & " )", PgString
        Next
    End Function

End Class
