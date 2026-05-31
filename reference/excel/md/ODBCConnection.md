# ODBCConnection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002448E-0000-0000-C000-000000000046}  

Represents the ODBC connection.

**Remarks:** An ODBC connection can be stored in an Excel workbook. When Excel opens the workbook, it creates an in-memory copy of the ODBC connection known as the ODBCConnection object. An ODBCConnection object contains information related to the connection, such as the name of the server to connect to and the name of the objects to be opened on that server. Optionally, the ODBCConnection object may also include authentication credential information, or a command that is to be passed to the server and executed (for example, a SELECT statement to be executed by SQL Server).

## Properties (20)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `BackgroundQuery As Boolean  (read/write)`  
  True if queries for the ODBC connection are performed asynchronously (in the background). Read/write Boolean.
- `CommandText As Variant  (read/write)`  
  Returns or sets the command string for the specified data source. Read/write Variant.
- `CommandType As XlCmdType  (read/write)`  
  Returns or sets one of the XlCmdType constants. Read/write XlCmdType.
- `Connection As Variant  (read/write)`  
  Returns or sets a string that contains ODBC settings that enable Microsoft Excel to connect to an ODBC data source. Read/write Variant.
- `EnableRefresh As Boolean  (read/write)`  
  True if the connection can be refreshed by the user. The default value is True. Read/write Boolean.
- `RefreshDate As Date  (read-only)`  
  Returns the date on which the ODBC connection was last refreshed. Read-only Date.
- `Refreshing As Boolean  (read-only)`  
  True if a background ODBC query is in progress for the specified ODBC connection. Read/write Boolean.
- `RefreshOnFileOpen As Boolean  (read/write)`  
  True if the connection is automatically updated each time the workbook is opened. The default value is False. Read/write Boolean.
- `RefreshPeriod As Long  (read/write)`  
  Returns or sets the number of minutes between refreshes. Read/write Long.
- `RobustConnect As XlRobustConnect  (read/write)`  
  Returns or sets how an ODBC connection connects to its data source. Read/write XlRobustConnect.
- `SavePassword As Boolean  (read/write)`  
  True if password information in an ODBC connection string is saved in the connection string. False if the password is removed. Read/write Boolean.
- `SourceConnectionFile As String  (read/write)`  
  Returns or sets a String indicating the Microsoft Office Data Connection file or similar file that was used to create the connection. Read/write.
- `SourceData As Variant  (read/write)`  
  Returns the data source for the ODBC connection, as shown in the table. Read/write Variant.
- `SourceDataFile As String  (read/write)`  
  Returns or sets a String indicating the source data file for an ODBC connection. Read/write.
- `ServerCredentialsMethod As XlCredentialsMethod  (read/write)`  
  Returns or sets the type of credentials that should be used for server authentication. Read/write XlCredentialsMethod.
- `ServerSSOApplicationID As String  (read/write)`  
  Returns or sets a single sign-on application (SSO) identifier that is used to do a lookup in the SSO database for credentials. Read/write String.
- `AlwaysUseConnectionFile As Boolean  (read/write)`  
  True if the connection file is always used to establish connection to the data source. Read/write Boolean.

## Methods (3)

- `CancelRefresh()`  
  Cancels all refresh operations in progress for the specified ODBC connection.
- `Refresh()`  
  Refreshes an ODBC connection.
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`  
  Saves the ODBC connection as a Microsoft Office Data Connection file.
    - `ODCFileName As String` (required): Location to save the file.
    - `Description As Variant` (optional): Description that will be saved in the file.
    - `Keywords As Variant` (optional): Space-separated keywords that can be used to search for this file.
