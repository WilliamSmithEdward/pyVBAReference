# QueryTable

**Type:** Class  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {59191DA1-EA47-11CE-A51F-00AA0061507F}  

## Properties (62)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Name As String  (read/write)`
- `FieldNames As Boolean  (read/write)`
- `RowNumbers As Boolean  (read/write)`
- `FillAdjacentFormulas As Boolean  (read/write)`
- `RefreshOnFileOpen As Boolean  (read/write)`
- `Refreshing As Boolean  (read-only)`
- `FetchedRowOverflow As Boolean  (read-only)`
- `BackgroundQuery As Boolean  (read/write)`
- `RefreshStyle As XlCellInsertionMode  (read/write)`
- `EnableRefresh As Boolean  (read/write)`
- `SavePassword As Boolean  (read/write)`
- `Destination As Range  (read-only)`
- `Connection As Variant  (read/write)`
- `PostText As String  (read/write)`
- `ResultRange As Range  (read-only)`
- `Parameters As Parameters  (read-only)`
- `Recordset As Object  (read/write)`
- `SaveData As Boolean  (read/write)`
- `EnableEditing As Boolean  (read/write)`
- `TextFilePlatform As Long  (read/write)`
- `TextFileStartRow As Long  (read/write)`
- `TextFileParseType As XlTextParsingType  (read/write)`
- `TextFileTextQualifier As XlTextQualifier  (read/write)`
- `TextFileConsecutiveDelimiter As Boolean  (read/write)`
- `TextFileTabDelimiter As Boolean  (read/write)`
- `TextFileSemicolonDelimiter As Boolean  (read/write)`
- `TextFileCommaDelimiter As Boolean  (read/write)`
- `TextFileSpaceDelimiter As Boolean  (read/write)`
- `TextFileOtherDelimiter As String  (read/write)`
- `TextFileColumnDataTypes As Variant  (read/write)`
- `TextFileFixedColumnWidths As Variant  (read/write)`
- `PreserveColumnInfo As Boolean  (read/write)`
- `PreserveFormatting As Boolean  (read/write)`
- `AdjustColumnWidth As Boolean  (read/write)`
- `CommandText As Variant  (read/write)`
- `CommandType As XlCmdType  (read/write)`
- `TextFilePromptOnRefresh As Boolean  (read/write)`
- `QueryType As XlQueryType  (read-only)`
- `MaintainConnection As Boolean  (read/write)`
- `TextFileDecimalSeparator As String  (read/write)`
- `TextFileThousandsSeparator As String  (read/write)`
- `RefreshPeriod As Long  (read/write)`
- `WebSelectionType As XlWebSelectionType  (read/write)`
- `WebFormatting As XlWebFormatting  (read/write)`
- `WebTables As String  (read/write)`
- `WebPreFormattedTextToColumns As Boolean  (read/write)`
- `WebSingleBlockTextImport As Boolean  (read/write)`
- `WebDisableDateRecognition As Boolean  (read/write)`
- `WebConsecutiveDelimitersAsOne As Boolean  (read/write)`
- `WebDisableRedirections As Boolean  (read/write)`
- `EditWebPage As Variant  (read/write)`
- `SourceConnectionFile As String  (read/write)`
- `SourceDataFile As String  (read/write)`
- `RobustConnect As XlRobustConnect  (read/write)`
- `TextFileTrailingMinusNumbers As Boolean  (read/write)`
- `ListObject As ListObject  (read-only)`
- `TextFileVisualLayout As XlTextVisualLayoutType  (read/write)`
- `WorkbookConnection As WorkbookConnection  (read-only)`
- `Sort As Sort  (read-only)`

## Methods (5)

- `CancelRefresh()`
- `Delete()`
- `Refresh([BackgroundQuery As Variant]) As Boolean`
- `ResetTimer()`
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`

## Events (2)

- `BeforeRefresh(Cancel As Boolean)`
- `AfterRefresh(Success As Boolean)`
