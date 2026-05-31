# PivotTable

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020872-0000-0000-C000-000000000046}  

Represents a PivotTable report on a worksheet.

**Remarks:** The PivotTable object is a member of the PivotTables collection. The PivotTables collection contains all the PivotTable objects on a single worksheet. Because PivotTable report programming can be complex, it's generally easiest to record PivotTable report actions and then revise the recorded code.

**Example:**

```vba
Worksheets("Sheet3").PivotTables(1) _
 .PivotFields("Year").Orientation = xlRowField
```

## Properties (105)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ColumnFields As Object  (read-only)`  
  Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently shown as column fields. Read-only.
- `ColumnGrand As Boolean  (read/write)`  
  True if the PivotTable report shows grand totals for columns. Read/write Boolean.
- `ColumnRange As Range  (read-only)`  
  Returns a Range object that represents the range that contains the column area in the PivotTable report. Read-only.
- `DataBodyRange As Range  (read-only)`  
  Returns a Range object that represents the range of values in a PivotTable. Read-only.
- `DataFields As Object  (read-only)`  
  Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently shown as data fields. Read-only.
- `DataLabelRange As Range  (read-only)`  
  Returns a Range object that represents the range that contains the labels for the data fields in the PivotTable report. Read-only.
- `_Default As String  (read/write)`
- `HasAutoFormat As Boolean  (read/write)`  
  True if the PivotTable report is automatically formatted when it's refreshed or when fields are moved. Read/write Boolean.
- `HiddenFields As Object  (read-only)`  
  Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently not shown as row, column, page, or data fields. Read-only.
- `InnerDetail As String  (read/write)`  
  Returns or sets the name of the field that will be shown as detail when the ShowDetail property is True for the innermost row or column field. Read/write String.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `PageFields As Object  (read-only)`  
  Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently showing as page fields. Read-only.
- `PageRange As Range  (read-only)`  
  Returns a Range object that represents the range that contains the page area in the PivotTable report. Read-only.
- `PageRangeCells As Range  (read-only)`  
  Returns a Range object that represents only the cells in the specified PivotTable report that contain the page fields and item drop-down lists.
- `RefreshDate As Date  (read-only)`  
  Returns the date on which the PivotTable report was last refreshed. Read-only Date.
- `RefreshName As String  (read-only)`  
  Returns the name of the person who last refreshed the PivotTable report data. Read-only String.
- `RowFields As Object  (read-only)`  
  Returns an object that represents either a single field in a PivotTable report (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently showing as row fields. Read-only.
- `RowGrand As Boolean  (read/write)`  
  True if the PivotTable report shows grand totals for rows. Read/write Boolean.
- `RowRange As Range  (read-only)`  
  Returns a Range object that represents the range including the row area on the PivotTable report. Read-only.
- `SaveData As Boolean  (read/write)`  
  True if data for the PivotTable report is saved with the workbook. False if only the report definition is saved. Read/write Boolean.
- `SourceData As Variant  (read/write)`  
  Returns the data source for the PivotTable report, as shown in the following table. Read/write Variant.
- `TableRange1 As Range  (read-only)`  
  Returns a Range object that represents the range containing the entire PivotTable report, but doesn't include page fields. Read-only.
- `TableRange2 As Range  (read-only)`  
  Returns a Range object that represents the range containing the entire PivotTable report, including page fields. Read-only.
- `Value As String  (read/write)`  
  Returns or sets a String value that represents the name of the PivotTable report.
- `VisibleFields As Object  (read-only)`  
  Returns an object that represents either a single field in a PivotTable report (a PivotField object) or a collection of all the visible fields (a PivotFields object). Visible fields are shown as row, column, page or data fields. Read-only.
- `CacheIndex As Long  (read/write)`  
  Returns or sets the index number of the PivotTable cache. Read/write Long.
- `DisplayErrorString As Boolean  (read/write)`  
  True if the PivotTable report displays a custom error string in cells that contain errors. The default value is False. Read/write Boolean.
- `DisplayNullString As Boolean  (read/write)`  
  True if the PivotTable report displays a custom string in cells that contain null values. The default value is True. Read/write Boolean.
- `EnableDrilldown As Boolean  (read/write)`  
  True if drilldown is enabled. The default value is True. Read/write Boolean.
- `EnableFieldDialog As Boolean  (read/write)`  
  True if the PivotTable Field dialog box is available when the user double-clicks the PivotTable field. The default value is True. Read/write Boolean.
- `EnableWizard As Boolean  (read/write)`  
  True if the PivotTable Wizard is available. The default value is True. Read/write Boolean.
- `ErrorString As String  (read/write)`  
  Returns or sets a String value that represents the string displayed in cells that contain errors when the DisplayErrorString property is True.
- `ManualUpdate As Boolean  (read/write)`  
  True if the PivotTable report is recalculated only at the user's request. The default value is False. Read/write Boolean.
- `MergeLabels As Boolean  (read/write)`  
  True if the specified PivotTable report's outer-row item, column item, subtotal, and grand total labels use merged cells. Read/write Boolean.
- `NullString As String  (read/write)`  
  Returns or sets the string displayed in cells that contain null values when the DisplayNullString property is True. The default value is an empty string (""). Read/write String.
- `PivotFormulas As PivotFormulas  (read-only)`  
  Returns a PivotFormulas object that represents the collection of formulas for the specified PivotTable report. Read-only.
- `SubtotalHiddenPageItems As Boolean  (read/write)`  
  True if hidden page field items in the PivotTable report are included in row and column subtotals, block totals, and grand totals. The default value is False. Read/write Boolean.
- `PageFieldOrder As Long  (read/write)`  
  Returns or sets the order in which page fields are added to the PivotTable report's layout. Can be one of the following XlOrder constants: xlDownThenOver or xlOverThenDown. The default constant is xlDownThenOver. Read/write Long.
- `PageFieldStyle As String  (read/write)`  
  Returns or sets the style used in the bound page field area. The default value is a null string (no style is applied by default). Read/write String.
- `PageFieldWrapCount As Long  (read/write)`  
  Returns or sets the number of page fields in each column or row in the PivotTable report. Read/write Long.
- `PreserveFormatting As Boolean  (read/write)`  
  True if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.
- `PivotSelection As String  (read/write)`  
  Returns or sets the PivotTable selection in standard PivotTable report selection format. Read/write String.
- `SelectionMode As XlPTSelectionMode  (read/write)`  
  Returns or sets the PivotTable report structured selection mode. Read/write XlPTSelectionMode.
- `Tag As String  (read/write)`  
  Returns or sets a string saved with the PivotTable report. Read/write String.
- `VacatedStyle As String  (read/write)`  
  Returns or sets the style applied to cells vacated when the PivotTable report is refreshed. The default value is a null string (no style is applied by default). Read/write String.
- `PrintTitles As Boolean  (read/write)`  
  True if the print titles for the worksheet are set based on the PivotTable report. False if the print titles for the worksheet are used. The default value is False. Read/write Boolean.
- `CubeFields As CubeFields  (read-only)`  
  Returns the CubeFields collection. Each CubeField object contains the properties of the cube field element. Read-only.
- `GrandTotalName As String  (read/write)`  
  Returns or sets the text string label that is displayed in the grand total column or row heading in the specified PivotTable report. The default value is the string Grand Total. Read/write String.
- `SmallGrid As Boolean  (read/write)`  
  True if Microsoft Excel uses a grid that's two cells wide and two cells deep for a newly created PivotTable report. False if Excel uses a blank stencil outline. Read/write Boolean.
- `RepeatItemsOnEachPrintedPage As Boolean  (read/write)`  
  True if row, column, and item labels appear on the first row of each page when the specified PivotTable report is printed. False if labels are printed only on the first page. The default value is True. Read/write Boolean.
- `TotalsAnnotation As Boolean  (read/write)`  
  True if an asterisk (\) is displayed next to each subtotal and grand total value in the specified PivotTable report if the report is based on an OLAP data source. The default value is True. Read/write Boolean**.
- `PivotSelectionStandard As String  (read/write)`  
  Returns or sets a String indicating the PivotTable selection in standard PivotTable report format using English (United States) settings. Read/write.
- `DataPivotField As PivotField  (read-only)`  
  Returns a PivotField object that represents all the data fields in a PivotTable. Read-only.
- `EnableDataValueEditing As Boolean  (read/write)`  
  True to disable the alert for when the user overwrites values in the data area of the PivotTable. True also allows the user to change data values that previously could not be changed. The default value is False. Read/write Boolean.
- `MDX As String  (read-only)`  
  Returns a String indicating the Multidimensional Expression (MDX) that would be sent to the provider to populate the current PivotTable view. Read-only.
- `ViewCalculatedMembers As Boolean  (read/write)`  
  When set to True (default), calculated members for Online Analytical Processing (OLAP) PivotTables can be viewed. Read/write Boolean.
- `CalculatedMembers As CalculatedMembers  (read-only)`  
  Returns a CalculatedMembers collection representing all the calculated members and calculated measures for an OLAP PivotTable.
- `DisplayImmediateItems As Boolean  (read/write)`  
  Returns or sets a Boolean that indicates whether items in the row and column areas are visible when the data area of the PivotTable is empty. Set this property to False to hide the items in the row and column areas when the data area of the PivotTable is empty. The default value is True.
- `EnableFieldList As Boolean  (read/write)`  
  False to disable the ability to display the field list for the PivotTable. If the field list was already being displayed, it disappears. The default value is True. Read/write Boolean.
- `VisualTotals As Boolean  (read/write)`  
  True (default) to enable Online Analytical Processing (OLAP) PivotTables to retotal after an item has been hidden from view. Read/write Boolean.
- `ShowPageMultipleItemLabel As Boolean  (read/write)`  
  When set to True (default), "(Multiple Items)" will appear in the PivotTable cell on the worksheet whenever items are hidden and an aggregate of non-hidden items is shown in the PivotTable view. Read/write Boolean.
- `Version As XlPivotTableVersionList  (read-only)`  
  Returns an XlPivotTableVersionList value that represents the Microsoft Excel version number.
- `DisplayEmptyRow As Boolean  (read/write)`  
  Returns True when the non-empty MDX keyword is included in the query to the OLAP provider for the category axis. The OLAP provider will not return empty rows in the result set. Returns False when the non-empty keyword is omitted. Read/write Boolean.
- `DisplayEmptyColumn As Boolean  (read/write)`  
  Returns True when the non-empty MDX keyword is included in the query to the OLAP provider for the value axis. The OLAP provider will not return empty columns in the result set. Returns False when the non-empty keyword is omitted. Read/write Boolean.
- `PivotColumnAxis As PivotAxis  (read-only)`  
  Returns a PivotAxis object representing the entire column axis. Read-only PivotAxis.
- `PivotRowAxis As PivotAxis  (read-only)`  
  Returns a PivotAxis object representing the entire row axis. Read-only PivotAxis.
- `ShowDrillIndicators As Boolean  (read/write)`  
  The ShowDrillIndicators property is used for toggling the display of drill indicators in the PivotTable. Read/write Boolean.
- `PrintDrillIndicators As Boolean  (read/write)`  
  Specifies whether drill indicators are printed with the PivotTable. Read/write Boolean.
- `DisplayMemberPropertyTooltips As Boolean  (read/write)`  
  Controls whether to display member properties in tooltips. Read/write Boolean.
- `DisplayContextTooltips As Boolean  (read/write)`  
  Controls whether tooltips are displayed for PivotTable cells. Read/write Boolean.
- `CompactRowIndent As Long  (read/write)`  
  Returns or sets the indent increment for PivotItems when compact row layout form is turned on. Read/write.
- `LayoutRowDefault As XlLayoutRowType  (read/write)`  
  This property specifies the layout settings for PivotFields when they are added to the PivotTable for the first time. Read/write XlLayoutRowType.
- `DisplayFieldCaptions As Boolean  (read/write)`  
  Controls whether filter buttons and PivotField captions for rows and columns are displayed in the grid. Read/write.
- `ActiveFilters As PivotFilters  (read-only)`  
  Indicates the currently active filter in the specified PivotTable. Read-only.
- `InGridDropZones As Boolean  (read/write)`  
  This property is used to toggle in-grid drop zones for a PivotTable object. In some cases, it also affects the layout of the PivotTable. Read/write Boolean.
- `TableStyle2 As Variant  (read/write)`  
  The TableStyle2 property specifies the PivotTable style currently applied to the PivotTable. Read/write.
- `ShowTableStyleLastColumn As Boolean  (read/write)`  
  Returns or sets if the last column is displayed for the specified PivotTable object. Read/write Boolean.
- `ShowTableStyleRowStripes As Boolean  (read/write)`  
  The ShowTableStyleRowStripes property displays banded rows in which even rows are formatted differently from odd rows. This makes PivotTables easier to read. Read/write Boolean.
- `ShowTableStyleColumnStripes As Boolean  (read/write)`  
  The ShowTableStyleColumnStripes property displays banded columns in which even columns are formatted differently from odd columns. This makes PivotTables easier to read. Read/write Boolean.
- `ShowTableStyleRowHeaders As Boolean  (read/write)`  
  The ShowTableStyleRowHeaders property is set to True if the row headers should be displayed in the PivotTable. Read/write Boolean.
- `ShowTableStyleColumnHeaders As Boolean  (read/write)`  
  The ShowTableStyleColumnHeaders property is set to True if the column headers should be displayed in the PivotTable. Read/write Boolean.
- `AllowMultipleFilters As Boolean  (read/write)`  
  Sets or retrieves a value that indicates whether a PivotField can have multiple filters applied to it at the same time. Read/write Boolean.
- `CompactLayoutRowHeader As String  (read/write)`  
  Specifies the caption that is displayed in the row header of a PivotTable when in compact row layout form. Read/write String.
- `CompactLayoutColumnHeader As String  (read/write)`  
  Specifies the caption that is displayed in the column header of a PivotTable when in compact row layout form. Read/write String.
- `FieldListSortAscending As Boolean  (read/write)`  
  Controls the sort order of fields in the PivotTable Field List. When this property is set to True, the fields are sorted in ascending order. When it is set to False, the fields are sorted in data source order. Read/write.
- `SortUsingCustomLists As Boolean  (read/write)`  
  The SortUsingCustomLists property controls whether custom lists are used for sorting items of fields, both initially when the PivotField is initialized and the PivotItems are ordered by their captions, and later when the user applies a sort. Read/write Boolean.
- `Location As String  (read/write)`  
  Gets or sets a String that represents the top-left cell in the body of the specified PivotTable object. Read/write.
- `EnableWriteback As Boolean  (read/write)`  
  Returns or sets whether writing back to the data source is enabled for the specified PivotTable. The default value is False. Read/write.
- `Allocation As XlAllocation  (read/write)`  
  Returns or sets whether to run an UPDATE CUBE statement for each cell that is edited, or only when the user chooses to calculate changes when performing what-if analysis on a PivotTable based on an OLAP data source. Read/write.
- `AllocationValue As XlAllocationValue  (read/write)`  
  Returns or sets the value to allocate when performing what-if analysis on a PivotTable report based on an OLAP data source. Read/write.
- `AllocationMethod As XlAllocationMethod  (read/write)`  
  Returns or sets the method to use to allocate values when performing what-if analysis on a PivotTable report based on an OLAP data source. Read/write.
- `AllocationWeightExpression As String  (read/write)`  
  Returns or sets the MDX weight expression to use when performing what-if analysis on a PivotTable report based on an OLAP data source. Read/write.
- `ChangeList As PivotTableChangeList  (read-only)`  
  Returns the PivotTableChangeList collection that represents the list of changes that have been made to the specified PivotTable based on an OLAP data source. Read-only.
- `Slicers As Slicers  (read-only)`  
  Returns the Slicers collection for the specified PivotTable. Read-only.
- `AlternativeText As String  (read/write)`  
  Returns or sets the descriptive (alternative) text string for the specified PivotTable. Read/write.
- `Summary As String  (read/write)`  
  Returns or sets the description associated with the alternative text string for the specified PivotTable. Read/write.
- `VisualTotalsForSets As Boolean  (read/write)`  
  Returns or sets whether to include filtered items in the totals of named sets for the specified PivotTable. Read/write.
- `ShowValuesRow As Boolean  (read/write)`  
  Returns or sets whether the values row is displayed. Read/write.
- `CalculatedMembersInFilters As Boolean  (read/write)`  
  Returns or sets whether to evaluate calculated members from OLAP servers in filters. Read/write.
- `Hidden As Boolean  (read-only)`  
  Checks whether the PivotTable exists at the worksheet level. Read-only Boolean.
- `PivotChart As Shape  (read-only)`  
  Returns a Shape object that represents the standalone PivotChart for the specified hidden PivotTable report. Read-only.
- `AutoRefresh As Boolean  (read/write)`

## Methods (31)

- `AddFields([RowFields As Variant], [ColumnFields As Variant], [PageFields As Variant], [AddToTable As Variant]) As Variant`  
  Adds row, column, and page fields to a PivotTable report or PivotChart report.
    - `RowFields As Variant` (optional): Specifies a field name (or an array of field names) to be added as rows or added to the category axis.
    - `ColumnFields As Variant` (optional): Specifies a field name (or an array of field names) to be added as columns or added to the series axis.
    - `PageFields As Variant` (optional): Specifies a field name (or an array of field names) to be added as pages or added to the page area.
    - `AddToTable As Variant` (optional): Applies only to PivotTable reports. True to add the specified fields to the report (none of the existing fields are replaced). False to replace existing fields with the new fields. The default value is False.
- `ShowPages([PageField As Variant]) As Variant`  
  Creates a new PivotTable report for each item in the page field. Each new report is created on a new worksheet.
    - `PageField As Variant` (optional): A string that names a single page field in the report.
- `PivotFields([Index As Variant]) As Object`  
  Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of both the visible and hidden fields (a PivotFields object) in the PivotTable report. Read-only.
    - `Index As Variant` (optional): The name or number of the field to be returned.
- `RefreshTable() As Boolean`  
  Refreshes the PivotTable report from the source data. Returns True if it's successful.
- `CalculatedFields() As CalculatedFields`  
  Returns a CalculatedFields collection that represents all the calculated fields in the specified PivotTable report. Read-only.
- `GetData(Name As String) As Double`  
  Returns the value for the data filed in a PivotTable.
    - `Name As String` (required): Describes a single cell in the PivotTable report, using syntax similar to the PivotSelect method or the PivotTable report references in calculated item formulas.
- `ListFormulas()`  
  Creates a list of calculated PivotTable items and fields on a separate worksheet.
- `PivotCache() As PivotCache`  
  Returns a PivotCache object that represents the cache for the specified PivotTable report. Read-only.
- `PivotTableWizard([SourceType As Variant], [SourceData As Variant], [TableDestination As Variant], [TableName As Variant], [RowGrand As Variant], [ColumnGrand As Variant], [SaveData As Variant], [HasAutoFormat As Variant], [AutoPage As Variant], [Reserved As Variant], [BackgroundQuery As Variant], [OptimizeCache As Variant], [PageFieldOrder As Variant], [PageFieldWrapCount As Variant], [ReadData As Variant], [Connection As Variant])`  
  Creates and returns a PivotTable object. This method doesn't display the PivotTable Wizard. This method isn't available for OLE DB data sources. Use the Add method to add a PivotTable cache, and then create a PivotTable report based on the cache.
    - `SourceType As Variant` (optional): An XlPivotTableSourceType value that represents the source of the report data. If you specify this argument, you must also specify _SourceData_. If _SourceType_ and _SourceData_ are omitted, Microsoft Excel assumes that the source type is xlDatabase, and the source data comes from the named range Database. If this named range doesn't exist, Excel uses the current region if the current selection is in a range of more than 10 cells that contain data. If this isn't true, this method will fail.
    - `SourceData As Variant` (optional): The data for the new report. Can be a Range object, an array of ranges, or a text constant that represents the name of another report. For an external database, _SourceData_ is an array of strings containing the SQL query string, where each element is up to 255 characters in length. You should use the _Connection_ argument to specify the ODBC connection string. For compatibility with earlier versions of Excel, _SourceData_ can be a two-element array. The first element is the connection string specifying the ODBC source for the data. The second element is the SQL query string used to get the data. If you specify _SourceData_, you must also specify _SourceType_. If the active cell is inside the _SourceData_ range, you must specify _TableDestination_ as well.
    - `TableDestination As Variant` (optional): A Range object specifying where the report should be placed on the worksheet. If this argument is omitted, the report is placed at the active cell.
    - `TableName As Variant` (optional): A string that specifies the name of the new report.
    - `RowGrand As Variant` (optional): True to show grand totals for rows in the report.
    - `ColumnGrand As Variant` (optional): True to show grand totals for columns in the report.
    - `SaveData As Variant` (optional): True to save data with the report. False to save only the report definition.
    - `HasAutoFormat As Variant` (optional): True to have Excel automatically format the report when it's refreshed or when fields are moved.
    - `AutoPage As Variant` (optional): Valid only if _SourceType_ is xlConsolidation. True to have Excel create a page field for the consolidation. If _AutoPage_ is False, you must create the page field or fields.
    - `Reserved As Variant` (optional): Not used by Excel.
    - `BackgroundQuery As Variant` (optional): True to have Excel perform queries for the report asynchronously (in the background). The default value is False.
    - `OptimizeCache As Variant` (optional): True to optimize the PivotTable cache when it's constructed. The default value is False.
    - `PageFieldOrder As Variant` (optional): The order in which page fields are added to the PivotTable report's layout. Can be one of the following XlOrder constants: xlDownThenOver (default) or xlOverThenDown.
    - `PageFieldWrapCount As Variant` (optional): The number of page fields in each column or row in the PivotTable report. The default value is 0 (zero).
    - `ReadData As Variant` (optional): True to create a PivotTable cache that contains all records from the external database; this cache can be very large. If _ReadData_ is False, you can set some of the fields as server-based page fields before the data is actually read.
    - `Connection As Variant` (optional): A string that contains ODBC settings that allow Excel to connect to an ODBC data source. The connection string has the form ODBC;<connection string>. This argument overrides any previous setting for the PivotCache object's Connection property.
- `Update()`  
  Updates the PivotTable report.
- `PivotSelect(Name As String, [Mode As XlPTSelectionMode], [UseStandardName As Variant])`  
  Selects part of a PivotTable report.
    - `Name As String` (required): The part of the PivotTable report to select.
    - `Mode As XlPTSelectionMode` (optional): Specifies the structured selection mode.
    - `UseStandardName As Variant` (optional): True for recorded macros that will play back in other locales.
- `GetPivotData([DataField As Variant], [Field1 As Variant], [Item1 As Variant], [Field2 As Variant], [Item2 As Variant], [Field3 As Variant], [Item3 As Variant], [Field4 As Variant], [Item4 As Variant], [Field5 As Variant], [Item5 As Variant], [Field6 As Variant], [Item6 As Variant], [Field7 As Variant], [Item7 As Variant], [Field8 As Variant], [Item8 As Variant], [Field9 As Variant], [Item9 As Variant], [Field10 As Variant], [Item10 As Variant], [Field11 As Variant], [Item11 As Variant], [Field12 As Variant], [Item12 As Variant], [Field13 As Variant], [Item13 As Variant], [Field14 As Variant], [Item14 As Variant]) As Range`  
  Returns a Range object with information about a data item in a PivotTable report.
    - `DataField As Variant` (optional): The name of the field containing the data for the PivotTable.
    - `Field1 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item1 As Variant` (optional): The name of an item in _Field1_.
    - `Field2 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item2 As Variant` (optional): The name of an item in _Field2_.
    - `Field3 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item3 As Variant` (optional): The name of an item in _Field3_.
    - `Field4 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item4 As Variant` (optional): The name of an item in _Field4_.
    - `Field5 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item5 As Variant` (optional): The name of an item in _Field5_.
    - `Field6 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item6 As Variant` (optional): The name of an item in _Field6_.
    - `Field7 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item7 As Variant` (optional): The name of an item in _Field7_.
    - `Field8 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item8 As Variant` (optional): The name of an item in _Field8_.
    - `Field9 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item9 As Variant` (optional): The name of an item in _Field9_.
    - `Field10 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item10 As Variant` (optional): The name of an item in _Field10_.
    - `Field11 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item11 As Variant` (optional): The name of an item in _Field11_.
    - `Field12 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item12 As Variant` (optional): The name of an item in _Field12_.
    - `Field13 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item13 As Variant` (optional): The name of an item in _Field13_.
    - `Field14 As Variant` (optional): The name of a column or row field in the PivotTable report.
    - `Item14 As Variant` (optional): The name of an item in _Field14_.
- `AddDataField(Field As Object, [Caption As Variant], [Function As Variant]) As PivotField`  
  Adds a data field to a PivotTable report. Returns a PivotField object that represents the new data field.
    - `Field As Object` (required): The unique field on the server. If the source data is Online Analytical Processing (OLAP), the unique field is a cube field. If the source data is non-OLAP (non-OLAP source data), the unique field is a PivotTable field.
    - `Caption As Variant` (optional): The label used in the PivotTable report to identify this data field.
    - `Function As Variant` (optional): The function performed in the added data field.
- `CreateCubeFile(File As String, [Measures As Variant], [Levels As Variant], [Members As Variant], [Properties As Variant]) As String`  
  Creates a cube file from a PivotTable report connected to an Online Analytical Processing (OLAP) data source.
    - `File As String` (required): The name of the cube file to be created. It will overwrite the file if it already exists.
    - `Measures As Variant` (optional): An array of unique names of measures that are to be part of the slice.
    - `Levels As Variant` (optional): An array of strings. Each array item is a unique level name. It represents the lowest level of a hierarchy that is in the slice.
    - `Members As Variant` (optional): An array of string arrays. The elements correspond, in order, to the hierarchies represented in the _Levels_ array. Each element is an array of string arrays that consists of the unique names of the top level members in the dimension that are to be included in the slice.
    - `Properties As Variant` (optional): False results in no member properties being included in the slice. The default value is True.
- `ClearTable()`  
  The ClearTable method is used for clearing a PivotTable. Clearing PivotTables includes removing all the fields and deleting all filtering and sorting applied to the PivotTables. This method resets the PivotTable to the state it had right after it was created, before any fields were added to it.
- `RowAxisLayout(RowLayout As XlLayoutRowType)`  
  This method is used for simultaneously setting layout options for all existing PivotFields.
    - `RowLayout As XlLayoutRowType` (required): Specifies the type of layout row. Can be xlCompactRow, xlTabularRow, or xlOutlineRow.
- `SubtotalLocation(Location As XlSubtototalLocationType)`  
  This method changes the subtotal location for all existing PivotFields. Changing the subtotal location has an immediate visual effect only for fields in outline form, but it will be set for fields in tabular form as well.
    - `Location As XlSubtototalLocationType` (required): XlSubtotalLocationType can be either xlAtTop or xlAtBottom.
- `ClearAllFilters()`  
  The ClearAllFilters method deletes all filters currently applied to the PivotTable. This includes deleting all filters in the PivotFilters collection, removing any manual filtering applied, and setting all PivotFields in the Report Filter area to the default item.
- `ConvertToFormulas(ConvertFilters As Boolean)`  
  The ConvertToFormulas method is used for converting a PivotTable to cube formulas. Read/write Boolean.
    - `ConvertFilters As Boolean` (required): Contains True or False to indicate the state of the ReportFilter area.
- `ChangeConnection(conn As WorkbookConnection)`  
  Changes the connection of the specified PivotTable object.
    - `conn As WorkbookConnection` (required): A WorkbookConnection object that represents the new connection for the PivotTable.
- `ChangePivotCache(PivotCache As Variant)`  
  Changes the PivotCache object of the specified PivotTable.
- `AllocateChanges()`  
  Performs a writeback operation for all edited cells in a PivotTable report based on an OLAP data source.
- `CommitChanges()`  
  Performs a commit operation on the data source of a PivotTable report based on an OLAP data source.
- `DiscardChanges()`  
  Discards all changes in the edited cells of a PivotTable report based on an OLAP data source.
- `RefreshDataSourceValues()`  
  Retrieves the current values from the data source for all edited cells in a PivotTable report that is in writeback mode.
- `RepeatAllLabels(Repeat As XlPivotFieldRepeatLabels)`  
  Specifies whether to repeat item labels for all PivotFields in the specified PivotTable.
    - `Repeat As XlPivotFieldRepeatLabels` (required): Specifies whether to repeat all field item labels in a PivotTable report.
- `PivotValueCell([rowline As Variant], [columnline As Variant]) As PivotValueCell`  
  Retrieve the PivotValueCell object for a given PivotTable provided certain row and column indices.
    - `rowline As Variant` (optional): If of type PivotLine, specifies the PivotLine in the row area that the PivotValueCell is aligned with. If of type Int, specifies the position of the PivotLine on the row area that the PivotValueCell is aligned with. If missing, Empty, Null, or 0 specifies the grand total row.
    - `columnline As Variant` (optional): If of type PivotLine, specifies the PivotLine in the column area that the PivotValueCell is aligned with. If of type Int, specifies the position of the PivotLine on the column area that the PivotValueCell is aligned with. If missing, Empty, Null or 0 specifies the grand total column.
- `DrillDown(PivotItem As PivotItem, [PivotLine As Variant])`  
  Enables you to drill down into the data within an OLAP-based or PowerPivot-based cube hierarchy.
    - `PivotItem As PivotItem` (required): The member from which the drill down is performed.
    - `PivotLine As Variant` (optional): Specifies the line in the PivotTable where the operation starting member resides. In cases where PivotLine is not specified, defaults to the top PivotLine where the member appears.
- `DrillUp(PivotItem As PivotItem, [PivotLine As Variant], [LevelUniqueName As Variant])`  
  Enables you to drill up into the data within an OLAP-based or PowerPivot-based cube hierarchy.
    - `PivotItem As PivotItem` (required): The member from which the drill up is performed.
    - `PivotLine As Variant` (optional): Specifies the line in the PivotTable where the operation starting member resides. In cases where PivotLine is not specified, defaults to the top PivotLine where the member appears.
    - `LevelUniqueName As Variant` (optional): The target for a multi-level drill up. The default action, if not specified, is a one level drill up.
- `DrillTo(PivotItem As PivotItem, CubeField As CubeField, [PivotLine As Variant])`  
  Enables you to drill to a location within an OLAP-based or PowerPivot-based cube hierarchy.
    - `PivotItem As PivotItem` (required): The member from which the drill operation is performed.
    - `CubeField As CubeField` (required): The target hierarchy being drilled to.
    - `PivotLine As Variant` (optional): Specifies the line in the PivotTable where the operation starting member resides. In cases where PivotLine is not specified, defaults to the top PivotLine where the member appears.
- `ApplyLayout()`
