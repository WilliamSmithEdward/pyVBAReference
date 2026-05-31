# PivotCache

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002441C-0000-0000-C000-000000000046}  

Represents the memory cache for a PivotTable report.

**Remarks:** The PivotCache object is a member of the PivotCaches collection.

**Example:**

```vba
Worksheets(1).PivotTables(1).PivotCache.RefreshOnFileOpen = True
```

## Properties (34)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `BackgroundQuery As Boolean  (read/write)`  
  True if queries for the PivotTable report are performed asynchronously (in the background). Read/write Boolean.
- `Connection As Variant  (read/write)`  
  Returns or sets a string that contains one of the following:
- `EnableRefresh As Boolean  (read/write)`  
  True if the PivotTable cache or query table can be refreshed by the user. The default value is True. Read/write Boolean.
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `MemoryUsed As Long  (read-only)`  
  Returns the amount of memory currently being used by the object, in bytes. Read-only Long.
- `OptimizeCache As Boolean  (read/write)`  
  True if the PivotTable cache is optimized when it's constructed. The default value is False. Read/write Boolean.
- `RecordCount As Long  (read-only)`  
  Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only Long.
- `RefreshDate As Date  (read-only)`  
  Returns the date on which the cache was last refreshed. Read-only Date.
- `RefreshName As String  (read-only)`  
  Returns the name of the person who last refreshed the PivotTable cache. Read-only String.
- `RefreshOnFileOpen As Boolean  (read/write)`  
  True if the PivotTable cache is automatically updated each time the workbook is opened. The default value is False. Read/write Boolean.
- `SavePassword As Boolean  (read/write)`  
  True if password information in an ODBC connection string is saved with the specified query. False if the password is removed. Read/write Boolean.
- `SourceData As Variant  (read/write)`  
  Returns the data source for the PivotTable report, as shown in the following table. Read/write Variant.
- `CommandText As Variant  (read/write)`  
  Returns or sets the command string for the specified data source. Read/write Variant.
- `CommandType As XlCmdType  (read/write)`  
  Returns or sets one of these XlCmdType constants: xlCmdCube, xlCmdDefault, xlCmdSql, or xlCmdTable.
- `QueryType As XlQueryType  (read-only)`  
  Indicates the type of query used by Microsoft Excel to populate the PivotTable cache. Read-only XlQueryType.
- `MaintainConnection As Boolean  (read/write)`  
  True if the connection to the specified data source is maintained after the refresh and until the workbook is closed. The default value is True. Read/write Boolean.
- `RefreshPeriod As Long  (read/write)`  
  Returns or sets the number of minutes between refreshes. Read/write Long.
- `Recordset As Object  (read/write)`  
  Returns or sets a Recordset object that's used as the data source for the specified PivotTable cache. Read/write.
- `LocalConnection As Variant  (read/write)`  
  Returns or sets the connection string to an offline cube file. Read/write String.
- `UseLocalConnection As Boolean  (read/write)`  
  Returns True if the LocalConnection property is used to specify the string that enables Microsoft Excel to connect to a data source. Returns False if the connection string specified by the Connection property is used. Read/write Boolean.
- `ADOConnection As Object  (read-only)`  
  Returns an ADO Connection object if the PivotTable cache is connected to an OLE DB data source. The ADOConnection property exposes the Microsoft Excel connection to the data provider, allowing the user to write code within the context of the same session that Excel is using with ADO (relational source) or ADO MD (OLAP source). Read-only.
- `IsConnected As Boolean  (read-only)`  
  Returns True if the MaintainConnection property is True, and the PivotTable cache is currently connected to its source. Returns False if it is not currently connected to its source. Read-only Boolean.
- `OLAP As Boolean  (read-only)`  
  Returns True if the PivotTable cache is connected to an Online Analytical Processing (OLAP) server. Read-only Boolean.
- `SourceType As XlPivotTableSourceType  (read-only)`  
  Returns an XlPivotTableSourceType value that represents the type of item being published.
- `MissingItemsLimit As XlPivotTableMissingItems  (read/write)`  
  Returns or sets the maximum quantity of unique items per PivotTable field that are retained even when they have no supporting data in the cache records. Read/write XlPivotTableMissingItems.
- `SourceConnectionFile As String  (read/write)`  
  Returns or sets a String indicating the Microsoft Office Data Connection file or similar file that was used to create the PivotTable. Read/write.
- `SourceDataFile As String  (read-only)`  
  Returns a String value that indicates the source data file for the cache of the PivotTable.
- `RobustConnect As XlRobustConnect  (read/write)`  
  Returns or sets how the PivotTable cache connects to its data source. Read/write XlRobustConnect.
- `WorkbookConnection As WorkbookConnection  (read-only)`  
  Establishes a connection between the current workbook and the PivotCache object. Read-only.
- `Version As XlPivotTableVersionList  (read-only)`  
  Returns the version of Microsoft Excel in which the PivotCache was created. Read-only XlPivotTableVersionList.
- `UpgradeOnRefresh As Boolean  (read/write)`  
  Contains information on whether to upgrade the PivotCache and all connected PivotTables on the next refresh. Read/write Boolean.

## Methods (6)

- `Refresh()`  
  Causes the specified PivotCache to be redrawn immediately.
- `ResetTimer()`  
  Resets the refresh timer for the specified query table or PivotTable report to the last interval that you set by using the RefreshPeriod property.
- `CreatePivotTable(TableDestination As Variant, [TableName As Variant], [ReadData As Variant], [DefaultVersion As Variant]) As PivotTable`  
  Creates a PivotTable report based on a PivotCache object. Returns a PivotTable object.
    - `TableDestination As Variant` (required): The cell in the upper-left corner of the PivotTable report's destination range (the range on the worksheet where the resulting PivotTable report will be placed). The destination range must be on a worksheet in the workbook that contains the PivotCache object specified by _expression_.
    - `TableName As Variant` (optional): The name of the new PivotTable report.
    - `ReadData As Variant` (optional): True to create a PivotTable cache that contains all the records from the external database; this cache can be very large. False to enable setting some of the fields as server-based page fields before the data is actually read.
    - `DefaultVersion As Variant` (optional): The default version of the PivotTable report.
- `MakeConnection()`  
  Establishes a connection for the specified PivotTable cache.
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`  
  Saves the PivotTable cache source as a Microsoft Office Data Connection file.
    - `ODCFileName As String` (required): Location to save the file.
    - `Description As Variant` (optional): Description that will be saved in the file.
    - `Keywords As Variant` (optional): Space-separated keywords that can be used to search for this file.
- `CreatePivotChart(ChartDestination As Variant, [XlChartType As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As Shape`  
  Creates a standalone PivotChart from a PivotCache object. Returns a Shape object.
    - `ChartDestination As Variant` (required): The Destination worksheet.
    - `XlChartType As Variant` (optional): The type of chart.
    - `Left As Variant` (optional): The distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).
    - `Top As Variant` (optional): The distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.
    - `Width As Variant` (optional): The width, in points, of the object.
    - `Height As Variant` (optional): The height, in points, of the object.
