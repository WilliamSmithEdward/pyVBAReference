# QueryTable

**Type:** Class  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {59191DA1-EA47-11CE-A51F-00AA0061507F}  

Represents a worksheet table built from data returned from an external data source, such as a SQL server or a Microsoft Access database.

**Remarks:** The QueryTable object is a member of the QueryTables collection.

**Example:**

```vba
Sheets("sheet1").QueryTables(1).FillAdjacentFormulas = True
```

## Properties (62)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `FieldNames As Boolean  (read/write)`  
  True if field names from the data source appear as column headings for the returned data. The default value is True. Read/write Boolean.
- `RowNumbers As Boolean  (read/write)`  
  True if row numbers are added as the first column of the specified query table. Read/write Boolean.
- `FillAdjacentFormulas As Boolean  (read/write)`  
  True if formulas to the right of the specified query table are automatically updated whenever the query table is refreshed. Read/write Boolean.
- `RefreshOnFileOpen As Boolean  (read/write)`  
  True if the PivotTable cache or query table is automatically updated each time the workbook is opened. The default value is False. Read/write Boolean.
- `Refreshing As Boolean  (read-only)`  
  True if there is a background query in progress for the specified query table. Read-only Boolean.
- `FetchedRowOverflow As Boolean  (read-only)`  
  True if the number of rows returned by the last use of the Refresh method is greater than the number of rows available on the worksheet. Read-only Boolean.
- `BackgroundQuery As Boolean  (read/write)`  
  True if queries for the query table are performed asynchronously (in the background). Read/write Boolean.
- `RefreshStyle As XlCellInsertionMode  (read/write)`  
  Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query. Read/write XlCellInsertionMode.
- `EnableRefresh As Boolean  (read/write)`  
  True if the PivotTable cache or query table can be refreshed by the user. The default value is True. Read/write Boolean.
- `SavePassword As Boolean  (read/write)`  
  True if password information in an ODBC connection string is saved with the specified query. False if the password is removed. Read/write Boolean.
- `Destination As Range  (read-only)`  
  Returns the cell in the upper-left corner of the query table destination range (the range where the resulting query table will be placed). The destination range must be on the worksheet that contains the QueryTable object. Read-only Range.
- `Connection As Variant  (read/write)`  
  Returns or sets a string that contains one of the following:
- `PostText As String  (read/write)`  
  Returns or sets the string used with the post method of inputting data into a web server to return data from a web query. Read/write String.
- `ResultRange As Range  (read-only)`  
  Returns a Range object that represents the area of the worksheet occupied by the specified query table. Read-only.
- `Parameters As Parameters  (read-only)`  
  Returns a Parameters collection that represents the query table parameters. Read-only.
- `Recordset As Object  (read/write)`  
  Returns or sets a Recordset object that's used as the data source for the specified query table. Read/write.
- `SaveData As Boolean  (read/write)`  
  True if data for the QueryTable report is saved with the workbook. False if only the report definition is saved. Read/write Boolean.
- `EnableEditing As Boolean  (read/write)`  
  True if the user can edit the specified query table. False if the user can only refresh the query table. Read/write Boolean.
- `TextFilePlatform As Long  (read/write)`  
  Returns or sets the origin of the text file that you are importing into the query table. This property determines which code page is used during the data import. Read/write XlPlatform.
- `TextFileStartRow As Long  (read/write)`  
  Returns or sets the row number at which text parsing will begin when you import a text file into a query table. Valid values are integers from 1 through 32767. The default value is 1. Read/write Long.
- `TextFileParseType As XlTextParsingType  (read/write)`  
  Returns or sets the column format for the data in the text file that you are importing into a query table. Read/write XlTextParsingType.
- `TextFileTextQualifier As XlTextQualifier  (read/write)`  
  Returns or sets the text qualifier when you import a text file into a query table. The text qualifier specifies that the enclosed data is in text format. Read/write XlTextQualifier.
- `TextFileConsecutiveDelimiter As Boolean  (read/write)`  
  True if consecutive delimiters are treated as a single delimiter when you import a text file into a query table. The default value is False. Read/write Boolean.
- `TextFileTabDelimiter As Boolean  (read/write)`  
  True if the tab character is the delimiter when you import a text file into a query table. The default value is False. Read/write Boolean.
- `TextFileSemicolonDelimiter As Boolean  (read/write)`  
  True if the semicolon is the delimiter when you import a text file into a query table, and if the value of the TextFileParseType property is xlDelimited. The default value is False. Read/write Boolean.
- `TextFileCommaDelimiter As Boolean  (read/write)`  
  True if the comma is the delimiter when you import a text file into a query table. False if you want to use some other character as the delimiter. The default value is False. Read/write Boolean.
- `TextFileSpaceDelimiter As Boolean  (read/write)`  
  True if the space character is the delimiter when you import a text file into a query table. The default value is False. Read/write Boolean.
- `TextFileOtherDelimiter As String  (read/write)`  
  Returns or sets the character used as the delimiter when you import a text file into a query table. The default value is null. Read/write String.
- `TextFileColumnDataTypes As Variant  (read/write)`  
  Returns or sets an ordered array of constants that specify the data types applied to the corresponding columns in the text file that you are importing into a query table. The default constant for each column is xlGeneral. Read/write Variant.
- `TextFileFixedColumnWidths As Variant  (read/write)`  
  Returns or sets an array of integers that correspond to the widths of the columns (in characters) in the text file that you are importing into a query table. Valid widths are from 1 through 32767 characters. Read/write Variant.
- `PreserveColumnInfo As Boolean  (read/write)`  
  True if column sorting, filtering, and layout information is preserved whenever a query table is refreshed. The default value is True. Read/write Boolean.
- `PreserveFormatting As Boolean  (read/write)`  
  True if any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren't formatted. The property is False if the last AutoFormat applied to the query table is applied to new rows of data. The default value is True.
- `AdjustColumnWidth As Boolean  (read/write)`  
  True if the column widths are automatically adjusted for the best fit each time you refresh the specified query table. False if the column widths are not automatically adjusted with each refresh. The default value is True. Read/write Boolean.
- `CommandText As Variant  (read/write)`  
  Returns or sets the command string for the specified data source. Read/write Variant.
- `CommandType As XlCmdType  (read/write)`  
  Returns or sets one of these XlCmdType constants: xlCmdCube, xlCmdDefault, xlCmdSql, or xlCmdTable. The constant that is returned or set describes the value of the CommandText property. The default value is xlCmdSQL. Read/write XlCmdType.
- `TextFilePromptOnRefresh As Boolean  (read/write)`  
  True if you want to specify the name of the imported text file each time the query table is refreshed. The Import Text File dialog box allows you to specify the path and file name. The default value is False. Read/write Boolean.
- `QueryType As XlQueryType  (read-only)`  
  Indicates the type of query used by Microsoft Excel to populate the query table. Read-only XlQueryType.
- `MaintainConnection As Boolean  (read/write)`  
  True if the connection to the specified data source is maintained after the refresh and until the workbook is closed. The default value is True. Read/write Boolean.
- `TextFileDecimalSeparator As String  (read/write)`  
  Returns or sets the decimal separator character that Microsoft Excel uses when you import a text file into a query table. The default is the system decimal separator character. Read/write String.
- `TextFileThousandsSeparator As String  (read/write)`  
  Returns or sets the thousands separator character that Microsoft Excel uses when you import a text file into a query table. The default is the system thousands separator character. Read/write String.
- `RefreshPeriod As Long  (read/write)`  
  Returns or sets the number of minutes between refreshes. Read/write Long.
- `WebSelectionType As XlWebSelectionType  (read/write)`  
  Returns or sets a value that determines whether an entire webpage, all tables on the webpage, or only specific tables on the webpage are imported into a query table. Read/write XlWebSelectionType.
- `WebFormatting As XlWebFormatting  (read/write)`  
  Returns or sets a value that determines how much formatting from a webpage, if any, is applied when you import the page into a query table. Read/write XlWebFormatting.
- `WebTables As String  (read/write)`  
  Returns or sets a comma-delimited list of table names or table index numbers when you import a webpage into a query table. Read/write String.
- `WebPreFormattedTextToColumns As Boolean  (read/write)`  
  Returns or sets whether data contained within HTML <PRE> tags on the webpage is parsed into columns when you import the page into a query table. The default is True. Read/write Boolean.
- `WebSingleBlockTextImport As Boolean  (read/write)`  
  True if data from the HTML <PRE> tags on the specified webpage is processed all at once when you import the page into a query table. False if the data is imported in blocks of contiguous rows so that header rows will be recognized as such. The default value is False. Read/write Boolean.
- `WebDisableDateRecognition As Boolean  (read/write)`  
  True if data that resembles dates is parsed as text when you import a webpage into a query table. False if date recognition is used. The default value is False. Read/write Boolean.
- `WebConsecutiveDelimitersAsOne As Boolean  (read/write)`  
  True if consecutive delimiters are treated as a single delimiter when you import data from HTML <PRE> tags on a webpage into a query table, and if the data is to be parsed into columns. False if you want to treat consecutive delimiters as multiple delimiters. The default value is True. Read/write Boolean.
- `WebDisableRedirections As Boolean  (read/write)`  
  True if web query redirections are disabled for a QueryTable object. The default value is False. Read/write Boolean.
- `EditWebPage As Variant  (read/write)`  
  Returns or sets the webpage Uniform Resource Locator (URL) for a web query. Read/write Variant.
- `SourceConnectionFile As String  (read/write)`  
  Returns or sets a String indicating the Microsoft Office Data Connection file or similar file that was used to create the QueryTable. Read/write.
- `SourceDataFile As String  (read/write)`  
  Returns or sets a String value that indicates the source data file for a query table.
- `RobustConnect As XlRobustConnect  (read/write)`  
  Returns or sets how the query table connects to its data source. Read/write XlRobustConnect.
- `TextFileTrailingMinusNumbers As Boolean  (read/write)`  
  True for Microsoft Excel to treat numbers imported as text that begin with a - (minus) symbol as a negative symbol. False for Excel to treat numbers imported as text that begin with a - symbol as text. Read/write Boolean.
- `ListObject As ListObject  (read-only)`  
  Returns a ListObject object for the QueryTable object. Read-only ListObject object.
- `TextFileVisualLayout As XlTextVisualLayoutType  (read/write)`  
  Returns or sets an XlTextVisualLayoutType enumeration that indicates whether the visual layout of the text being imported is left-to-right or right-to-left.
- `WorkbookConnection As WorkbookConnection  (read-only)`  
  Returns the WorkbookConnection object that the query table uses. Read-only.
- `Sort As Sort  (read-only)`  
  Returns the sort criteria for the query table range. Read-only.

## Methods (5)

- `CancelRefresh()`  
  Cancels all background queries for the specified query table. Use the Refreshing property to determine whether a background query is currently in progress.
- `Delete()`  
  Deletes the object.
- `Refresh([BackgroundQuery As Variant]) As Boolean`  
  Updates an external data range in a QueryTable object.
    - `BackgroundQuery As Variant` (optional): Used only with QueryTables that are based on the results of a SQL query. True to return control to the procedure as soon as a database connection is made and the query is submitted. The QueryTable is updated in the background. False to return control to the procedure only after all data has been fetched to the worksheet. If this argument isn't specified, the setting of the BackgroundQuery property determines the query mode.
- `ResetTimer()`  
  Resets the refresh timer for the specified query table or PivotTable report to the last interval that you set by using the RefreshPeriod property.
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`  
  Saves the QueryTable cache source as a Microsoft Office Data Connection file.
    - `ODCFileName As String` (required): Location to save the file.
    - `Description As Variant` (optional): Description that will be saved in the file.
    - `Keywords As Variant` (optional): Space-separated keywords that can be used to search for this file.

## Events (2)

- `BeforeRefresh(Cancel As Boolean)`  
  Occurs before any refreshes of the query table. This includes refreshes resulting from calling the Refresh method, from the user's actions in the product, and from opening the workbook containing the query table.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the refresh doesn't occur when the procedure is finished.
- `AfterRefresh(Success As Boolean)`  
  Occurs after a query is completed or canceled.
    - `Success As Boolean` (required): True if the query was completed successfully.
