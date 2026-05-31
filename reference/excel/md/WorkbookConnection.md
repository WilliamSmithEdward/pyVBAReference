# WorkbookConnection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024485-0000-0000-C000-000000000046}  

A connection is a set of information needed to obtain data from an external data source other than a Microsoft Excel workbook.

**Remarks:** Connections can be stored within an Excel workbook. When the workbook is opened, Excel creates an in-memory copy of the connection that is referred to as the connection object. A connection object contains information such as the name of the server and the name of the object to be opened on that server. Optionally, the connection object may also include authentication credentials and/or a command that is to be passed to the server and executed (example: a SELECT statement to be executed by SQL Server). A connection may also be stored in a separate connection file. Most connections in an Excel workbook include a pointer to an external connection file. Connection files have extensions that clearly label them as connection files (.ODC, .IQY, etc.) and may be located on the user's local machine or in other well-known or trusted locations such as WSS (Data Connection Library), or other corporate servers. Connection files enable multiple users within the same organization to re-use connections. Network administrators are able to change the way the entire organization connects to a back-end data source by changing a single connection file.

## Properties (17)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read/write)`  
  Returns or sets the name of the WorkbookConnection object. Read/write String.
- `Description As String  (read/write)`  
  Returns or sets a brief description for a WorkbookConnection object. Read/write String.
- `_Default As String  (read/write)`
- `Type As XlConnectionType  (read-only)`  
  Returns the workbook connection type. Read-only XlConnectionType.
- `OLEDBConnection As OLEDBConnection  (read-only)`  
  Returns the OLEDB connection details for the specified WorkbookConnection object. Read-only OLEDBConnection.
- `ODBCConnection As ODBCConnection  (read-only)`  
  Returns the ODBC connection details for the specified WorkbookConnection object. Read-only ODBCConnection.
- `Ranges As Ranges  (read-only)`  
  Returns the range of objects for the specified WorkbookConnection object. Read-only Ranges.
- `ModelConnection As ModelConnection  (read-only)`  
  Returns an object that contains information for the new model connection type introduced in Excel 2013 to interact with the integrated Data Model. Read-only.
- `WorksheetDataConnection As WorksheetDataConnection  (read-only)`  
  Returns an object that contains information for a connection from the PowerPivot Model to data within the workbook, such as a range, named range, or table. Read-only.
- `RefreshWithRefreshAll As Boolean  (read/write)`  
  Determines if the connection should be refreshed when Refresh All is executed. Read/write Boolean.
- `TextConnection As TextConnection  (read-only)`  
  Returns a TextConnection object that contains the information on a query to a text file. Read-only.
- `DataFeedConnection As DataFeedConnection  (read-only)`  
  Returns a DataFeedConnection object that contains the data and functionality needed to connect to data feeds. Read-only.
- `InModel As Boolean  (read-only)`  
  Specifies whether the WorkbookConnection has been added to the model. Read-only Boolean.
- `ModelTables As ModelTables  (read-only)`  
  Returns a ModelTables object associated with the particular connection. Read-only.

## Methods (2)

- `Delete()`  
  Deletes a workbook connection.
- `Refresh()`  
  Refreshes a workbook connection.
