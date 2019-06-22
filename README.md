# Metric
This is agentless monitoring of Windows based systems. Used technologies such as WMI for accessing metric data and PostgreSQL database for storing results. For data visualization it is recommended to use the [Grafana project](https://github.com/grafana/grafana).


### How to run

Start cscript.exe with named parameters, like as:
```bash
cscript "metric.vbs /task:collect /metric:cpu /host:%local%
```
Parameters:
* /task
collect - run infinite loop for monitoring metrics
reinstall - recreate all tables for storage metric data

* /metric
You write name of metric. App instance will monitoring only one metric. List of metric names you find in files, which contains substring "plugin". This is class file of VBScript, which implements interface for collect and write metric data.

* /host
This parameter is optional and filtering host name in table rules. You can write name with wildcard chars. Only suitable for host name filter will be used to poll metrics.

### Plugins

Each plugin has vbscript code in file and implement functions and property:
* Property Name
[ to specify the name of the metric ]

* Function NewClass
[ this is constructor and returning class instance ]

* Function RecreateTable(PgString) 
 [ running only in reinstall task ]

* Function InsertRows(HostName, Params, PgString, WmiString)
[ main function for collect metric and write in PostgreSQL database. PgString and WmiString has connection strings for using in other Core Functions. You write only bussines logic, for saving and loading data function already available for using ]

Register module in "metric.vbs"
```vb
Set ModClass = IncludeModule("YourPluginName.vbs")
If Not ModClass Is Nothing Then ModuleList.Add ModClass.Name, ModClass
Set ModClass = Nothing
```

Thats all, you now have all the functions available in the Core.


### Core Functions

* Function toDateUniq
[ Returns a unique current date with milliseconds for use in sql queries. Each function call produces a unique value ]

* Function toDate(Value)
[ Format datetime value and return string for using in sql query ]

* Function toStr(Value)
[ Return string with single quoting ]

* Function toReal(Value)
[ Convert number with cyrilic decimal point for using in sql query]

* Function SelectPgData(Query, ConnectString)
[ Execute sql query and return array of rows. First dimesion has index of column, second dimesion has a row index. Use UBound(YourArray, 2) for detect count of rows ]

* Function ModifyPgData(Query, ConnectString)
[ Execute query and nothing to return. Use for DDL queries ]

* Function OpenPgConnection(ConnectString)
[ Function is provide lowest access for Connection Object. Use when you need transaction functionality. You decide when the connection will be closed and a commit or rollback is required ]

* Function ExecuteOnPgConnection(Query, ConnectionObject)
[ Executing query in Connection. Connection is not closing after executing. Use in transactions ]

* Function SelectWmiData(Query, ConnectString)
[ Function execute WQL query and return array of rows. Each item in array has a dictionary object. Key is name of column and Value is data of column ]

* Function SelectTaskParameters
[ Function parse command line and return dictionary of parameters. This is wrapper for append other logics, when parameter is not assigned and set default value  ]
