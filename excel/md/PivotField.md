# PivotField

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020874-0000-0000-C000-000000000046}  

Represents a field in a PivotTable report.

**Remarks:** The PivotField object is a member of the PivotFields collection. The PivotFields collection contains all the fields in a PivotTable report, including hidden fields. In some cases, it may be easier to use one of the properties that returns a subset of the PivotTable fields. The following properties are available: - ColumnFields property - DataFields property - HiddenFields property - PageFields property - RowFields property - VisibleFields property

**Example:**

```vba
Worksheets("sheet3").PivotTables(1) _
 .PivotFields("year").Orientation = xlRowField
```

## Properties (78)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Calculation As XlPivotFieldCalculation  (read/write)`  
  Returns or sets an XlPivotFieldCalculation value that represents the type of calculation performed by the specified field. This property is valid only for data fields.
- `ChildField As PivotField  (read-only)`  
  Returns a PivotField object that represents the child field for the specified field (if the field is grouped and has a child field). Read-only.
- `ChildItems As Variant  (read-only)`  
  Returns an object that represents either a single PivotTable item (a PivotItem object) or a collection of all the items (a PivotItems object) that are group children in the specified field, or children of the specified item. Read-only.
- `CurrentPage As Variant  (read/write)`  
  Returns or sets the current page showing for the page field (valid only for page fields). Read/write PivotItem.
- `DataRange As Range  (read-only)`  
  Returns a Range object as shown in the following table. Read-only.
- `DataType As XlPivotFieldDataType  (read-only)`  
  Returns an XlPivotFieldDataType value that represents the type of data in the PivotTable field.
- `_Default As String  (read/write)`
- `Function As XlConsolidationFunction  (read/write)`  
  Returns or sets the function used to summarize the PivotTable field (data fields only). Read/write XlConsolidationFunction.
- `GroupLevel As Variant  (read-only)`  
  Returns the placement of the specified field within a group of fields (if the field is a member of a grouped set of fields). Read-only.
- `HiddenItems As Variant  (read-only)`  
  Returns an object that represents either a single hidden PivotTable item (a PivotItem object) or a collection of all the hidden items (a PivotItems object) in the specified field. Read-only.
- `LabelRange As Range  (read-only)`  
  Returns a Range object that represents the cell (or cells) that contain the field label. Read-only.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `NumberFormat As String  (read/write)`  
  Returns or sets a String value that represents the format code for the object.
- `Orientation As XlPivotFieldOrientation  (read/write)`  
  Returns or sets an XlPivotFieldOrientation value that represents the location of the field in the specified PivotTable report.
- `ShowAllItems As Boolean  (read/write)`  
  True if all items in the PivotTable report are displayed, even if they don't contain summary data. The default value is False. Read/write Boolean.
- `ParentField As PivotField  (read-only)`  
  Returns a PivotField object that represents the PivotTable field that's the group parent of the specified object. The field must be grouped and must have a parent field. Read-only.
- `ParentItems As Variant  (read-only)`  
  Returns an object that represents either a single PivotTable item (a PivotItem object) or a collection of all the items (a PivotItems object) that are group parents in the specified field. The specified field must be a group parent of another field. Read-only.
- `Position As Variant  (read/write)`  
  Returns or sets a Variant value that represents the position of the field (first, second, third, and so on) among all the fields in its orientation (Rows, Columns, Pages, Data).
- `SourceName As String  (read-only)`  
  Returns a String value that represents the specified object's name as it appears in the original source data for the specified PivotTable report.
- `Subtotals As Variant  (read/write)`  
  Returns or sets subtotals displayed with the specified field. Valid only for nondata fields. Read/write Variant.
- `BaseField As Variant  (read/write)`  
  Returns or sets the base field for a custom calculation. This property is valid only for data fields. Read/write Variant.
- `BaseItem As Variant  (read/write)`  
  Returns or sets the item in the base field for a custom calculation. Valid only for data fields. Read/write Variant.
- `TotalLevels As Variant  (read-only)`  
  Returns the total number of fields in the current field group. If the field isn't grouped, or if the data source is OLAP-based, TotalLevels returns the value 1. Read-only Long.
- `Value As String  (read/write)`  
  Returns or sets a String value that represents the name of the specified field in the PivotTable report.
- `VisibleItems As Variant  (read-only)`  
  Returns an object that represents either a single visible PivotTable item (a PivotItem object) or a collection of all the visible items (a PivotItems object) in the specified field. Read-only.
- `DragToColumn As Boolean  (read/write)`  
  True if the specified field can be dragged to the column position. The default value is True. Read/write Boolean.
- `DragToHide As Boolean  (read/write)`  
  True if the field can be hidden by being dragged off the PivotTable report. The default value is True. Read/write Boolean.
- `DragToPage As Boolean  (read/write)`  
  True if the field can be dragged to the page position. The default value is True. Read/write Boolean.
- `DragToRow As Boolean  (read/write)`  
  True if the field can be dragged to the row position. The default value is True. Read/write Boolean.
- `DragToData As Boolean  (read/write)`  
  True if the specified field can be dragged to the data position. The default value is True. Read/write Boolean.
- `Formula As String  (read/write)`  
  Returns or sets a String value that represents the object's formula in A1-style notation and in the language of the macro.
- `IsCalculated As Boolean  (read-only)`  
  True if the PivotTable field is a calculated field or item. Read-only Boolean.
- `MemoryUsed As Long  (read-only)`  
  Returns the amount of memory currently being used by the object, in bytes. Read-only Long.
- `ServerBased As Boolean  (read/write)`  
  True if the data source for the specified PivotTable report is external and only the items matching the page field selection are retrieved. Read/write Boolean.
- `AutoSortOrder As Long  (read-only)`  
  Returns the order used to sort the specified PivotTable field automatically. Can be one of the XlSortOrder constants. Read-only Long.
- `AutoSortField As String  (read-only)`  
  Returns the name of the data field used to sort the specified PivotTable field automatically. Read-only String.
- `AutoShowType As Long  (read-only)`  
  Returns the xlAutomatic constant if AutoShow is enabled for the specified PivotTable field; returns xlManual if AutoShow is disabled. Read-only Long.
- `AutoShowRange As Long  (read-only)`  
  Returns the xlTop constant if the top items are shown automatically in the specified PivotTable field; returns xlBottom if the bottom items are shown. Read-only Long.
- `AutoShowCount As Long  (read-only)`  
  Returns the number of top or bottom items that are automatically shown in the specified PivotTable field. Read-only Long.
- `AutoShowField As String  (read-only)`  
  Returns the name of the data field used to determine the top or bottom items that are automatically shown in the specified PivotTable field. Read-only String.
- `LayoutBlankLine As Boolean  (read/write)`  
  True if a blank row is inserted after the specified row field in a PivotTable report. The default value is False. Read/write Boolean.
- `LayoutSubtotalLocation As XlSubtototalLocationType  (read/write)`  
  Returns or sets the position of the PivotTable field subtotals in relation to (either above or below) the specified field. Read/write XlSubtotalLocationType.
- `LayoutPageBreak As Boolean  (read/write)`  
  True if a page break is inserted after each field. The default value is False. Read/write Boolean.
- `LayoutForm As XlLayoutFormType  (read/write)`  
  Returns or sets the way the specified PivotTable items appear-in table format or in outline format. Read/write XlLayoutFormType.
- `SubtotalName As String  (read/write)`  
  Returns or sets the text string label displayed in the subtotal column or row heading in the specified PivotTable report. The default value is the string Subtotal. Read/write String.
- `Caption As String  (read/write)`  
  Returns a String value that represents the label text for the pivot field.
- `DrilledDown As Boolean  (read/write)`  
  True if the flag for the specified PivotTable field or PivotTable item is set to "drilled" (expanded, or visible). Read/write Boolean.
- `CubeField As CubeField  (read-only)`  
  Returns the CubeField object from which the specified PivotTable field is descended. Read-only.
- `CurrentPageName As String  (read/write)`  
  Returns or sets the currently displayed page of the specified PivotTable report. The name of the page appears in the page field. Note that this property works only if the currently displayed page already exists. Read/write String.
- `StandardFormula As String  (read/write)`  
  Returns or sets a String specifying formulas with standard English (United States) formatting. Read/write.
- `HiddenItemsList As Variant  (read/write)`  
  Returns or sets a Variant specifying an array of strings that are hidden items for a PivotTable field. Read/write.
- `DatabaseSort As Boolean  (read/write)`  
  When set to True, manual repositioning of items in a PivotTable field is allowed. Returns True if the field has no manually positioned items. Read/write Boolean.
- `IsMemberProperty As Boolean  (read-only)`  
  Returns True when the PivotField contains member properties. Read-only Boolean.
- `PropertyParentField As PivotField  (read-only)`  
  Returns a PivotField object representing the field to which the properties in this field pertain.
- `PropertyOrder As Long  (read/write)`  
  Valid only for PivotTable fields that are member property fields. Returns a Long indicating the display position of the member property within the cube field to which it belongs. Read/write.
- `EnableItemSelection As Boolean  (read/write)`  
  When set to False, disables the ability to use the field dropdown in the user interface. The default value is True. Read/write Boolean.
- `CurrentPageList As Variant  (read/write)`  
  Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report. Read/write Variant.
- `Hidden As Boolean  (read/write)`  
  This property is used to hide the individual levels of an OLAP hierarchy. Read/write Boolean.
- `UseMemberPropertyAsCaption As Boolean  (read/write)`  
  This property is used to control whether member property captions are used for PivotItem captions of the PivotField. Read/write Boolean.
- `MemberPropertyCaption As String  (read/write)`  
  Setting the MemberPropertyCaption property controls which member property is used as a caption for a given level. Read/write Boolean.
- `DisplayAsTooltip As Boolean  (read/write)`  
  This property is used to specify whether a specific member property PivotField is displayed in tooltips. Read/write Boolean.
- `DisplayInReport As Boolean  (read/write)`  
  This property is used to specify whether the specified member property PivotField is displayed in the PivotTable or not. Read/write Boolean.
- `DisplayAsCaption As Boolean  (read-only)`  
  This property is used to display member properties of PivotFields as captions. Read-only.
- `LayoutCompactRow As Boolean  (read/write)`  
  Specifies whether or not a PivotField is compacted (items of multiple PivotFields are displayed in a single column) when rows are selected. Read/write Boolean.
- `IncludeNewItemsInFilter As Boolean  (read/write)`  
  Allows developers to specify whether excluded or included items should be tracked when manual filtering is applied to the PivotField. Read/write Boolean.
- `VisibleItemsList As Variant  (read/write)`  
  Returns or sets a Variant specifying an array of strings that represent included items in a manual filter applied to a PivotField. Read/write.
- `PivotFilters As PivotFilters  (read-only)`  
  Returns or sets the PivotFilters for the specified PivotField object. Read-only.
- `AutoSortPivotLine As PivotLine  (read-only)`  
  Returns the name of the PivotLine object used to sort the specified PivotTable field automatically. Read-only.
- `AutoSortCustomSubtotal As Long  (read-only)`  
  Returns the name of the custom subtotal used to sort the specified PivotTable field automatically. Read-only.
- `ShowingInAxis As Boolean  (read-only)`  
  Indicates if the PivotField is currently visible in the PivotTable or not. Read-only.
- `EnableMultiplePageItems As Boolean  (read/write)`  
  Used for specifying whether check boxes are present in the filter drop-down list for fields in the page area. Read/write Boolean.
- `AllItemsVisible As Boolean  (read-only)`  
  Used to retrieve a Boolean value that indicates whether any manual filtering is applied to the PivotField. Read-only.
- `SourceCaption As String  (read-only)`  
  The SourceCaption property is applicable only for OLAP PivotTables, and returns the original caption from the OLAP server for a PivotField. Read-only.
- `ShowDetail As Boolean  (read/write)`  
  Gets or sets whether the specified PivotField object is showing detail. Read/write Boolean.
- `RepeatLabels As Boolean  (read/write)`  
  Returns or sets whether item labels are repeated in the PivotTable for the specified PivotField. Read/write.

## Methods (12)

- `PivotItems([Index As Variant]) As Variant`  
  Returns an object that represents either a single PivotTable item (a PivotItem object) or a collection of all the visible and hidden items (a PivotItems object) in the specified field. Read-only.
    - `Index As Variant` (optional): The name or number of the item to be returned.
- `CalculatedItems() As CalculatedItems`  
  Returns a CalculatedItems collection that represents all the calculated items in the specified PivotTable report. Read-only.
- `Delete()`  
  Deletes the object.
- `AutoShow(Type As Long, Range As Long, Count As Long, Field As String)`  
  Displays the number of top or bottom items for a row, page, or column field in the specified PivotTable report.
    - `Type As Long` (required): Use the xlAutomatic constant to cause the specified PivotTable report to show the items that match the specified criteria. Use xlManual to disable this feature.
    - `Range As Long` (required): The location at which to start showing items. Can be either of the following constants: xlTop or xlBottom.
    - `Count As Long` (required): The number of items to be shown.
    - `Field As String` (required): The name of the base data field. You must specify the unique name (as returned from the SourceName property), and not the displayed name.
- `AddPageItem(Item As String, [ClearList As Variant])`  
  Adds an additional item to a multiple item page field.
    - `Item As String` (required): Source name of a PivotItem object, corresponding to the specific Online Analytical Processing (OLAP) member unique name.
    - `ClearList As Variant` (optional): If False (default), adds a page item to the existing list. If True, deletes all current items and adds _Item_.
- `DrillTo(Field As String)`  
  The DrillTo method supports drilling to a specified PivotField from another PivotField.
- `ClearManualFilter()`  
  Provides an easy way to set the Visible property to True for all items of a PivotField in PivotTables, and to empty the HiddenItemsList and VisibleItemsList collections in OLAP PivotTables.
- `ClearAllFilters()`  
  Calling this method deletes all filters currently applied to the PivotField. This includes deleting all filters from the PivotFilters collection of the PivotField and removing any manual filtering applied to the PivotField as well. If the PivotField is in the Report Filter area, the item selected will be set to the default item.
- `ClearValueFilters()`  
  Calling this method deletes all value filters in the PivotFilters collection of the PivotField.
- `ClearLabelFilters()`  
  This method deletes all label filters or all date filters in the PivotFilters collection of the PivotField.
- `AutoSort(Order As Long, Field As String, [PivotLine As Variant], [CustomSubtotal As Variant])`  
  Establishes automatic field-sorting rules for PivotTable reports.
    - `Order As Long` (required): One of the XlSortOrder constants specifying the sort order.
    - `Field As String` (required): The name of the sort key field. You must specify the unique name (as returned from the SourceName property), and not the displayed name.
    - `PivotLine As Variant` (optional): A line on a column or row in a PivotTable report.
    - `CustomSubtotal As Variant` (optional): The custom subtotal field.
- `AutoGroup()`  
  Automatically groups the pivot fields in a PivotTable.
