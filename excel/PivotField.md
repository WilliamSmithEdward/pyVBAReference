# PivotField

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020874-0000-0000-C000-000000000046}  

## Properties (78)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Calculation As XlPivotFieldCalculation  (read/write)`
- `ChildField As PivotField  (read-only)`
- `ChildItems As Variant  (read-only)`
- `CurrentPage As Variant  (read/write)`
- `DataRange As Range  (read-only)`
- `DataType As XlPivotFieldDataType  (read-only)`
- `_Default As String  (read/write)`
- `Function As XlConsolidationFunction  (read/write)`
- `GroupLevel As Variant  (read-only)`
- `HiddenItems As Variant  (read-only)`
- `LabelRange As Range  (read-only)`
- `Name As String  (read/write)`
- `NumberFormat As String  (read/write)`
- `Orientation As XlPivotFieldOrientation  (read/write)`
- `ShowAllItems As Boolean  (read/write)`
- `ParentField As PivotField  (read-only)`
- `ParentItems As Variant  (read-only)`
- `Position As Variant  (read/write)`
- `SourceName As String  (read-only)`
- `Subtotals As Variant  (read/write)`
- `BaseField As Variant  (read/write)`
- `BaseItem As Variant  (read/write)`
- `TotalLevels As Variant  (read-only)`
- `Value As String  (read/write)`
- `VisibleItems As Variant  (read-only)`
- `DragToColumn As Boolean  (read/write)`
- `DragToHide As Boolean  (read/write)`
- `DragToPage As Boolean  (read/write)`
- `DragToRow As Boolean  (read/write)`
- `DragToData As Boolean  (read/write)`
- `Formula As String  (read/write)`
- `IsCalculated As Boolean  (read-only)`
- `MemoryUsed As Long  (read-only)`
- `ServerBased As Boolean  (read/write)`
- `AutoSortOrder As Long  (read-only)`
- `AutoSortField As String  (read-only)`
- `AutoShowType As Long  (read-only)`
- `AutoShowRange As Long  (read-only)`
- `AutoShowCount As Long  (read-only)`
- `AutoShowField As String  (read-only)`
- `LayoutBlankLine As Boolean  (read/write)`
- `LayoutSubtotalLocation As XlSubtototalLocationType  (read/write)`
- `LayoutPageBreak As Boolean  (read/write)`
- `LayoutForm As XlLayoutFormType  (read/write)`
- `SubtotalName As String  (read/write)`
- `Caption As String  (read/write)`
- `DrilledDown As Boolean  (read/write)`
- `CubeField As CubeField  (read-only)`
- `CurrentPageName As String  (read/write)`
- `StandardFormula As String  (read/write)`
- `HiddenItemsList As Variant  (read/write)`
- `DatabaseSort As Boolean  (read/write)`
- `IsMemberProperty As Boolean  (read-only)`
- `PropertyParentField As PivotField  (read-only)`
- `PropertyOrder As Long  (read/write)`
- `EnableItemSelection As Boolean  (read/write)`
- `CurrentPageList As Variant  (read/write)`
- `Hidden As Boolean  (read/write)`
- `UseMemberPropertyAsCaption As Boolean  (read/write)`
- `MemberPropertyCaption As String  (read/write)`
- `DisplayAsTooltip As Boolean  (read/write)`
- `DisplayInReport As Boolean  (read/write)`
- `DisplayAsCaption As Boolean  (read-only)`
- `LayoutCompactRow As Boolean  (read/write)`
- `IncludeNewItemsInFilter As Boolean  (read/write)`
- `VisibleItemsList As Variant  (read/write)`
- `PivotFilters As PivotFilters  (read-only)`
- `AutoSortPivotLine As PivotLine  (read-only)`
- `AutoSortCustomSubtotal As Long  (read-only)`
- `ShowingInAxis As Boolean  (read-only)`
- `EnableMultiplePageItems As Boolean  (read/write)`
- `AllItemsVisible As Boolean  (read-only)`
- `SourceCaption As String  (read-only)`
- `ShowDetail As Boolean  (read/write)`
- `RepeatLabels As Boolean  (read/write)`

## Methods (12)

- `PivotItems([Index As Variant]) As Variant`
- `CalculatedItems() As CalculatedItems`
- `Delete()`
- `AutoShow(Type As Long, Range As Long, Count As Long, Field As String)`
- `AddPageItem(Item As String, [ClearList As Variant])`
- `DrillTo(Field As String)`
- `ClearManualFilter()`
- `ClearAllFilters()`
- `ClearValueFilters()`
- `ClearLabelFilters()`
- `AutoSort(Order As Long, Field As String, [PivotLine As Variant], [CustomSubtotal As Variant])`
- `AutoGroup()`
