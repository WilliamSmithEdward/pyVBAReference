# OLEDBConnection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002448D-0000-0000-C000-000000000046}  

Represents the OLE DB connection.

**Remarks:** An OLE DB connection can be stored in an Excel workbook. When Excel opens the workbook, it creates an in-memory copy of the OLE DB connection known as the OLEDBConnection object. An OLEDBConnection object contains information related to the connection, such as the name of the server to connect to and the name of the objects to be opened on that server. Optionally, the OLEDBConnection object may also include authentication credential information, or a command that is to be passed to the server and executed (for example, a SELECT statement to be executed by SQL Server).

## Properties (33)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ADOConnection As Object  (read-only)`  
  Returns an ADO connection object if the PivotTable cache is connected to an OLE DB data source. Read-only.
- `BackgroundQuery As Boolean  (read/write)`  
  True if queries for the OLE DB connection are performed asynchronously (in the background). Read/write Boolean.
- `CommandText As Variant  (read/write)`  
  Returns or sets the command string for the specified data source. Read/write Variant.
- `CommandType As XlCmdType  (read/write)`  
  Returns or sets one of the XlCmdType constants. Read/write XlCmdType.
- `Connection As Variant  (read/write)`  
  Returns or sets a string that contains OLE DB settings that enable Microsoft Excel to connect to an OLE DB data source. Read/write Variant.
- `EnableRefresh As Boolean  (read/write)`  
  True if the connection can be refreshed by the user. The default value is True. Read/write Boolean.
- `LocalConnection As Variant  (read/write)`  
  Returns or sets the connection string to an offline cube file. Read/write String.
- `MaintainConnection As Boolean  (read/write)`  
  Returns True if the connection to the specified data source is maintained after the refresh operation and until the workbook is closed. Read/write Boolean.
- `RefreshDate As Date  (read-only)`  
  Returns the date on which the OLE DB connection was last refreshed. Read-only Date.
- `Refreshing As Boolean  (read-only)`  
  True if a background OLE DB query is in progress for the specified OLE DB connection. Read/write Boolean.
- `RefreshOnFileOpen As Boolean  (read/write)`  
  True if the connection is automatically updated each time the workbook is opened. The default value is False. Read/write Boolean.
- `RefreshPeriod As Long  (read/write)`  
  Returns or sets the number of minutes between refreshes. Read/write Long.
- `RobustConnect As XlRobustConnect  (read/write)`  
  Returns or sets how an OLE DB connection connects to its data source. Read/write XlRobustConnect.
- `SavePassword As Boolean  (read/write)`  
  True if password information in an OLE DB connection string is saved in the connection string. False if the password is removed. Read/write Boolean.
- `SourceConnectionFile As String  (read/write)`  
  Returns or sets a String indicating the Microsoft Office Data Connection file or similar file that was used to create the connection. Read/write.
- `SourceDataFile As String  (read/write)`  
  Returns or sets a String indicating the source data file for an OLE DB connection. Read/write.
- `OLAP As Boolean  (read-only)`  
  Returns True if the OLE DB connection is connected to an Online Analytical Processing (OLAP) server. Read-only Boolean.
- `UseLocalConnection As Boolean  (read/write)`  
  True if the LocalConnection property is used to specify the string that enables Microsoft Excel to connect to a data source. False if the connection string specified by the Connection property is used. Read/write Boolean.
- `MaxDrillthroughRecords As Long  (read/write)`  
  Returns or sets the maximum number of records to retrieve. Read/write Long.
- `IsConnected As Boolean  (read-only)`  
  Returns True if the MaintainConnection property is True. Returns False if it is not currently connected to its source. Read-only Boolean.
- `ServerCredentialsMethod As XlCredentialsMethod  (read/write)`  
  Returns or sets the type of credentials that should be used for server authentication. Read/write XlCredentialsMethod.
- `ServerSSOApplicationID As String  (read/write)`  
  Returns or sets a single sign-on application (SSO) identifier that is used to perform a lookup in the SSO database for credentials. Read/write String.
- `AlwaysUseConnectionFile As Boolean  (read/write)`  
  True if the connection file is always used to establish a connection to the data source. Read/write Boolean.
- `ServerFillColor As Boolean  (read/write)`  
  True if the fill color format for the OLAP server is retrieved from the server when using the specified connection. Read/write Boolean.
- `ServerFontStyle As Boolean  (read/write)`  
  True if the font style format for the OLAP server is retrieved from the server when using the specified connection. Read/write Boolean.
- `ServerNumberFormat As Boolean  (read/write)`  
  True if the number format for the OLAP server is retrieved from the server when using the specified connection. Read/write Boolean.
- `ServerTextColor As Boolean  (read/write)`  
  True if the text color format for the OLAP server is retrieved from the server when using the specified connection. Read/write Boolean.
- `RetrieveInOfficeUILang As Boolean  (read/write)`  
  True if the data and errors are to be retrieved in the Office user interface display language when available. Read/write Boolean.
- `CalculatedMembers As CalculatedMembers  (read-only)`  
  Returns the CalculatedMembers collection for the specified connection. Read-only.
- `LocaleID As Long  (read/write)`  
  Returns or sets the locale identifier for the specified connection. Read/write.

## Methods (5)

- `CancelRefresh()`  
  Cancels all refresh operations in progress for the specified OLE DB connection.
- `MakeConnection()`  
  Establishes a connection for the specified OLE DB connection.
- `Refresh()`  
  Refreshes an OLE DB connection.
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`  
  Saves the OLE DB connection as a Microsoft Office Data Connection file.
    - `ODCFileName As String` (required): Location to save the file.
    - `Description As Variant` (optional): Description that will be saved in the file.
    - `Keywords As Variant` (optional): Space-separated keywords that can be used to search for this file.
- `Reconnect()`  
  Drops and then reconnects the specified connection.
