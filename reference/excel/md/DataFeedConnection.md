# DataFeedConnection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244D4-0000-0000-C000-000000000046}  

Contains the data and functionality needed to connect to data feeds. The same object is used for all Data Feed types.

## Properties (16)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified DataFeedConnection object. Read-only.
- `AlwaysUseConnectionFile As Boolean  (read/write)`  
  True if the connection file is always used to establish a connection to the data source. Read/write Boolean.
- `CommandText As Variant  (read/write)`  
  Returns or sets the command string for the specified data source. Read/write Variant.
- `CommandType As XlCmdType  (read/write)`  
  Returns or sets the command string for the specified data source. Read/write Variant.
- `Connection As Variant  (read/write)`  
  Returns or sets a string that contains Service Contract settings that enable Microsoft Excel to connect to a Data Feed data source. Read/write Variant.
- `EnableRefresh As Boolean  (read/write)`  
  True if the connection can be refreshed by the user. The default value is True. Read/write Boolean.
- `RefreshDate As Date  (read-only)`  
  Returns the date on which the OLE DB connection was last refreshed. Read-only Date.
- `Refreshing As Boolean  (read-only)`  
  True if an OLE DB query is in progress for the specified data source connection. Read/write Boolean.
- `RefreshOnFileOpen As Boolean  (read/write)`  
  True if the connection is automatically updated each time the workbook is opened. The default value is False.
- `RefreshPeriod As Long  (read/write)`  
  Returns or sets the number of minutes between refreshes. Read/write Long.
- `SavePassword As Boolean  (read/write)`  
  True if password information in a data feed connection string is saved in the connection string. False if the password is removed.
- `ServerCredentialsMethod As XlCredentialsMethod  (read/write)`  
  Returns or sets the type of credentials that should be used for server authentication. Read/write XlCredentialsMethod enumeration.
- `SourceConnectionFile As String  (read/write)`  
  Returns or sets a String indicating the Microsoft Office Data Connection file or similar file that was used to create the connection. Read/write.
- `SourceDataFile As String  (read/write)`  
  A path to the original file used to create the connection. In the case of an OData connection, this is the location of the .atom or .atomsvc file used to create the connection. Read/write String.

## Methods (3)

- `CancelRefresh()`  
  Cancels a refresh operation on a data feed connection.
- `Refresh()`  
  Refreshes the data feed connection.
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`  
  Saves the data feed connection as a Microsoft Office Data Connection file.
    - `ODCFileName As String` (required): Location to save the file.
    - `Description As Variant` (optional): Description that will be saved in the file.
    - `Keywords As Variant` (optional): Space-separated keywords that can be used to search for this file.
