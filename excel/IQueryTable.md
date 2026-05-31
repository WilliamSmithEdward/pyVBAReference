# IQueryTable

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024428-0001-0000-C000-000000000046}  

## Properties (62)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Name As HRESULT  (read/write)`
- `FieldNames As HRESULT  (read/write)`
- `RowNumbers As HRESULT  (read/write)`
- `FillAdjacentFormulas As HRESULT  (read/write)`
- `RefreshOnFileOpen As HRESULT  (read/write)`
- `Refreshing As HRESULT  (read-only)`
- `FetchedRowOverflow As HRESULT  (read-only)`
- `BackgroundQuery As HRESULT  (read/write)`
- `RefreshStyle As HRESULT  (read/write)`
- `EnableRefresh As HRESULT  (read/write)`
- `SavePassword As HRESULT  (read/write)`
- `Destination As HRESULT  (read-only)`
- `Connection As HRESULT  (read/write)`
- `PostText As HRESULT  (read/write)`
- `ResultRange As HRESULT  (read-only)`
- `Parameters As HRESULT  (read-only)`
- `Recordset As HRESULT  (read/write)`
- `SaveData As HRESULT  (read/write)`
- `EnableEditing As HRESULT  (read/write)`
- `TextFilePlatform As HRESULT  (read/write)`
- `TextFileStartRow As HRESULT  (read/write)`
- `TextFileParseType As HRESULT  (read/write)`
- `TextFileTextQualifier As HRESULT  (read/write)`
- `TextFileConsecutiveDelimiter As HRESULT  (read/write)`
- `TextFileTabDelimiter As HRESULT  (read/write)`
- `TextFileSemicolonDelimiter As HRESULT  (read/write)`
- `TextFileCommaDelimiter As HRESULT  (read/write)`
- `TextFileSpaceDelimiter As HRESULT  (read/write)`
- `TextFileOtherDelimiter As HRESULT  (read/write)`
- `TextFileColumnDataTypes As HRESULT  (read/write)`
- `TextFileFixedColumnWidths As HRESULT  (read/write)`
- `PreserveColumnInfo As HRESULT  (read/write)`
- `PreserveFormatting As HRESULT  (read/write)`
- `AdjustColumnWidth As HRESULT  (read/write)`
- `CommandText As HRESULT  (read/write)`
- `CommandType As HRESULT  (read/write)`
- `TextFilePromptOnRefresh As HRESULT  (read/write)`
- `QueryType As HRESULT  (read-only)`
- `MaintainConnection As HRESULT  (read/write)`
- `TextFileDecimalSeparator As HRESULT  (read/write)`
- `TextFileThousandsSeparator As HRESULT  (read/write)`
- `RefreshPeriod As HRESULT  (read/write)`
- `WebSelectionType As HRESULT  (read/write)`
- `WebFormatting As HRESULT  (read/write)`
- `WebTables As HRESULT  (read/write)`
- `WebPreFormattedTextToColumns As HRESULT  (read/write)`
- `WebSingleBlockTextImport As HRESULT  (read/write)`
- `WebDisableDateRecognition As HRESULT  (read/write)`
- `WebConsecutiveDelimitersAsOne As HRESULT  (read/write)`
- `WebDisableRedirections As HRESULT  (read/write)`
- `EditWebPage As HRESULT  (read/write)`
- `SourceConnectionFile As HRESULT  (read/write)`
- `SourceDataFile As HRESULT  (read/write)`
- `RobustConnect As HRESULT  (read/write)`
- `TextFileTrailingMinusNumbers As HRESULT  (read/write)`
- `ListObject As HRESULT  (read-only)`
- `TextFileVisualLayout As HRESULT  (read/write)`
- `WorkbookConnection As HRESULT  (read-only)`
- `Sort As HRESULT  (read-only)`

## Methods (5)

- `CancelRefresh()`
- `Delete()`
- `Refresh([BackgroundQuery As Variant], RHS As Boolean)`
- `ResetTimer()`
- `SaveAsODC(ODCFileName As String, [Description As Variant], [Keywords As Variant])`
