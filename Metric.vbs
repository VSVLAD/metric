Option Explicit


' ������  ���������� ��� ������  ������

Dim ModuleList, ModClass
Set ModuleList = CreateObject("Scripting.Dictionary")

IncludeModule "Metric.Core.vbs"

Set ModClass = IncludeModule("Metric.Plugin.Cpu.vbs")
If Not ModClass Is Nothing Then ModuleList.Add ModClass.Name, ModClass
Set ModClass = Nothing

Set ModClass = IncludeModule("Metric.Plugin.Disk.vbs")
If Not ModClass Is Nothing Then ModuleList.Add ModClass.Name, ModClass
Set ModClass = Nothing

Set ModClass = IncludeModule("Metric.Plugin.Service.vbs")
If Not ModClass Is Nothing Then ModuleList.Add ModClass.Name, ModClass
Set ModClass = Nothing

Set ModClass = IncludeModule("Metric.Plugin.Process.vbs")
If Not ModClass Is Nothing Then ModuleList.Add ModClass.Name, ModClass
Set ModClass = Nothing

Set ModClass = IncludeModule("Metric.Plugin.NginxLog.vbs")
If Not ModClass Is Nothing Then ModuleList.Add ModClass.Name, ModClass
Set ModClass = Nothing


' ������  �������� ���������  ������

' ������ ���������� � PostgreSQL
Dim PgString: PgString = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=PostgreSQL35W;DATABASE=YOUR_DATABASE;SERVER=localhost;PORT=5432;UID=postgres;Password=YOUR_PASSWORD;SSLmode=disable;ReadOnly=0;Protocol=7.4;FakeOidIndex=0;ShowOidColumn=0;RowVersioning=0;ShowSystemTables=0;Fetch=100;UnknownSizes=0;MaxVarcharSize=255;MaxLongVarcharSize=8190;Debug=0;CommLog=0;UseDeclareFetch=0;TextAsLongVarchar=1;UnknownsAsLongVarchar=0;BoolsAsChar=1;Parse=0;ExtraSysTablePrefixes=;LFConversion=1;UpdatableCursors=1;TrueIsMinus1=0;BI=0;ByteaAsLongVarBinary=1;UseServerSidePrepare=1;LowerCaseIdentifier=0;D6=-101;XaOpt=1"""

' ������ ���������� � WMI
Dim WmiString: WmiString = "winmgmts:\\%host_name%\root\CIMV2"

' ����� ������, ����� ��������� ���������
Dim timeOutSleep: timeOutSleep = 5000

' ������ ���������� ����������
Dim colArgs: Set colArgs = SelectTaskParameters

' �������� �� ���� WSH
If Instr(1, LCase(WScript.FullName), "wscript.exe") >= 1 Then
    LogMessage "����������� CScript ��� ����� ������, ������ WScript. " & vbCrLf & _
               "������ ������: cscript ""metric.vbs"" /task:collect /metric:cpu /host:%local%"
    WScript.Quit
End If


' ������  ���������� �����  ������

Select Case colArgs("task")
    Case "collect"

        ' �������� �� ���������� ������
        If colArgs("metric") = "" Or Instr(1, colArgs("metric"), "%") >= 1 Then
            LogMessage "�� ������� �������, ������� ����� ��������! ���������� �������� ������� ����� �������� � ���������� ��������� ������"
            WScript.Quit
        End If

        LogMessage "�������� ���� ������ [" & colArgs("metric") & ", " & colArgs("host") & "]"

        ' ������� ����
        Do While True
            Call CollectMetric(colArgs("host"), colArgs("metric"), PgString, WmiString)
            Call WScript.Sleep(timeOutSleep)
        Loop

    Case "reinstall"

        ' ������������ ������ ������ � ������� ������
        Call ReinstallMetricRule(PgString)

        LogMessage "��������� ������������ ����� ������ ������"

        Dim moduleName
        For Each moduleName In ModuleList
            ModuleList(moduleName).RecreateTable(PgString)
            LogMessage "��������� ������������ ����� ��� ������: " & ModuleList(moduleName).Name
        Next

    Case Else

        LogMessage "�� ������� ���������� ������! ������� �����: collect, reinstall"

End Select


' ��������� ����� ������
Public Sub CollectMetric(HostName, MetricName, PgString, WmiString)
    Dim pg_delete_metrics, pg_select_metrics, resMetrics, idx, lastTimer

    pg_delete_metrics = Replace( _
                    " delete from metric_%metric_name%                                                        " & _
                    "       where id in ( select m.id                                                         " & _
                    "                       from metric_%metric_name% m,                                      " & _
                    "                            metric_rule mr                                               " & _
                    "                      where mr.host_name = m.host_name                                   " & _
                    "                        and mr.metric_name = mr.metric_name                              " & _
                    "                        and m.metric_date < now() - interval '1 second' * mr.life_time ) ", "%metric_name%", MetricName)

    pg_select_metrics = Replace(Replace( _
                    " select mr.host_name, mr.metric_name, mr.params                                                " & _
                    " from metric_rule mr                                                                           " & _
                    " left join ( select m.host_name, m.metric_name, max(m.metric_date) max_metric_date             " & _
                    "             from metric_%metric_name% m                                                       " & _
                    "             group by m.host_name, m.metric_name                                               " & _
                    "   ) mmax on mmax.host_name = mr.host_name                                                     " & _
                    "         and mmax.metric_name = mr.metric_name                                                 " & _
                    " where coalesce(mmax.max_metric_date, '2000-01-01') + interval '1 second' * mr.period <= now() " & _
                    "   and mr.metric_name = '%metric_name%'                                                        " & _
                    "   and mr.host_name like '%host_name%'                                                         ", "%metric_name%", MetricName), "%host_name%", HostName)

    ' ������� ������ ������� �������� ������
    Call ModifyPgData(pg_delete_metrics, PgString)

    ' ������� �������, ������� ��������� � ����������
    resMetrics = SelectPgData(pg_select_metrics, PgString)

    If UBound(resMetrics) > 0 Then
        Dim localHostName, localMetricName, localParams, localWmiString
        For idx = 0 To UBound(resMetrics, 2)

            ' ��������� ������� ����������
            localHostName   = resMetrics(0, idx)
            localMetricName = resMetrics(1, idx)
            localParams     = resMetrics(2, idx)
            localWmiString  = Replace(WmiString, "%host_name%", localHostName)

            ' ��������� ���������� ���������� �������
            If ModuleList.Exists(localMetricName) Then
                lastTimer = Timer
                Call ModuleList(MetricName).InsertRows(localHostName, localParams, PgString, localWmiString)

                Call LogMessage("������� [" & localHostName & ", " & localMetricName & "] ��������� �� " & Round(Timer - lastTimer, 2) & " ���")
            Else
                Call LogMessage("������� [" & localHostName & ", " & localMetricName & "] �� ���������. ����������� ������")
            End If
        Next
    End If
End Sub



' ����������� �������
Public Function LogMessage(Message)
    Dim lastErrNumber, lastErrDescription
    lastErrNumber = Err.Number
    lastErrDescription = Err.Description

    On Error Resume Next

    Dim fso, objFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objFile = fso.OpenTextFile("Metric.log", 8, True)

    If lastErrNumber <> 0 Then
        Call objFile.WriteLine(toDate(Now) & " " & Message & " ::Err [" & lastErrNumber & ", " & lastErrDescription & "]")
        Call WScript.Echo(toDate(Now)      & " " & Message & " ::Err [" & lastErrNumber & ", " & lastErrDescription & "]")
    Else
        Call objFile.WriteLine(toDate(Now) & " " & Message)
        Call WScript.Echo(toDate(Now)      & " " & Message)
    End If

    objFile.Close
End Function

' ������� ��������� ��������
Public Function IncludeModule(FilePath)
    On Error Resume Next

	Dim fso, parentFolderName
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    parentFolderName = fso.GetParentFolderName(WScript.ScriptFullName)
    FilePath = fso.BuildPath(ParentFolderName, FilePath)
    
    ExecuteGlobal fso.OpenTextFile(FilePath, 1).ReadAll

    If Err.Number <> 0 Then
        LogMessage "::Metric.IncludeModule.ExecuteGlobal [" & FilePath & "]"
        Set IncludeModule = Nothing
    Else
        Set IncludeModule = Eval("NewClass")
    End If
    
    On Error GoTo 0
End Function
