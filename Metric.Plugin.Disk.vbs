Option Explicit

' Обязательная функция для иницилазации класса
Public Function NewClass
    Set NewClass = New MetricDisk
End Function

Class MetricDisk

    ' Название метрики
    Public Property Get Name
        Name = "disk"
    End Property
    
    ' Пересоздание схемы
    Public Function RecreateTable(PgString)
        ModifyPgData " drop table if exists metric_disk;                             " & _
                     " create table metric_disk(id bigserial,                        " & _
                     "                          host_name varchar(50) not null,      " & _
                     "                          metric_name varchar(50) not null,    " & _
                     "                          metric_date timestamp default now(), " & _
                     "                          Caption varchar(100),                " & _
                     "                          VolumeName varchar(100),             " & _
                     "                          Size bigint,                         " & _
                     "                          FreeSpace bigint,                    " & _
                     "                          FileSystem varchar(50),              " & _
                     "                          DriveType integer)                   ", PgString
                     
    End Function

    ' Добавление свежих метрик
    Public Function InsertRows(HostName, Params, PgString, WmiString)
        Dim item, colItems, timestampNow

        colItems = SelectWmiData(" select Caption, VolumeName, Size, FreeSpace, FileSystem, DriveType from Win32_LogicalDisk where DriveType = 3 ", WmiString)
        timestampNow = toDate(Now)

        For Each item In colItems
            ModifyPgData " insert into metric_disk(host_name, metric_name, metric_date, Caption, VolumeName, Size, FreeSpace, FileSystem, DriveType) values (" & _
                                toStr(HostName) 		                                                                                                & " ," & _
                                toStr("disk") 		                                                                                                    & " ," & _
                                ToStr(timestampNow)		                                                                                                & " ," & _
                                ToStr(item("Caption"))		                                                                                            & " ," & _
                                ToStr(item("VolumeName")) 		                                                                                        & " ," & _
                                ToReal(item("Size")) 			                                                                                        & " ," & _
                                ToReal(item("FreeSpace"))                                                                                               & " ," & _
                                ToStr(item("FileSystem"))                                                                                               & " ," & _
                                ToReal(item("DriveType"))                                                                                               & " )", PgString
        Next
    End Function

End Class