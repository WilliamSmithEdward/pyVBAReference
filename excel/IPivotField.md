# IPivotField

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020874-0001-0000-C000-000000000046}  

## Properties (78)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Calculation As HRESULT  (read/write)`
- `ChildField As HRESULT  (read-only)`
- `ChildItems As HRESULT  (read-only)`
- `CurrentPage As HRESULT  (read/write)`
- `DataRange As HRESULT  (read-only)`
- `DataType As HRESULT  (read-only)`
- `_Default As HRESULT  (read/write)`
- `Function As HRESULT  (read/write)`
- `GroupLevel As HRESULT  (read-only)`
- `HiddenItems As HRESULT  (read-only)`
- `LabelRange As HRESULT  (read-only)`
- `Name As HRESULT  (read/write)`
- `NumberFormat As HRESULT  (read/write)`
- `Orientation As HRESULT  (read/write)`
- `ShowAllItems As HRESULT  (read/write)`
- `ParentField As HRESULT  (read-only)`
- `ParentItems As HRESULT  (read-only)`
- `Position As HRESULT  (read/write)`
- `SourceName As HRESULT  (read-only)`
- `Subtotals As HRESULT  (read/write)`
- `BaseField As HRESULT  (read/write)`
- `BaseItem As HRESULT  (read/write)`
- `TotalLevels As HRESULT  (read-only)`
- `Value As HRESULT  (read/write)`
- `VisibleItems As HRESULT  (read-only)`
- `DragToColumn As HRESULT  (read/write)`
- `DragToHide As HRESULT  (read/write)`
- `DragToPage As HRESULT  (read/write)`
- `DragToRow As HRESULT  (read/write)`
- `DragToData As HRESULT  (read/write)`
- `Formula As HRESULT  (read/write)`
- `IsCalculated As HRESULT  (read-only)`
- `MemoryUsed As HRESULT  (read-only)`
- `ServerBased As HRESULT  (read/write)`
- `AutoSortOrder As HRESULT  (read-only)`
- `AutoSortField As HRESULT  (read-only)`
- `AutoShowType As HRESULT  (read-only)`
- `AutoShowRange As HRESULT  (read-only)`
- `AutoShowCount As HRESULT  (read-only)`
- `AutoShowField As HRESULT  (read-only)`
- `LayoutBlankLine As HRESULT  (read/write)`
- `LayoutSubtotalLocation As HRESULT  (read/write)`
- `LayoutPageBreak As HRESULT  (read/write)`
- `LayoutForm As HRESULT  (read/write)`
- `SubtotalName As HRESULT  (read/write)`
- `Caption As HRESULT  (read/write)`
- `DrilledDown As HRESULT  (read/write)`
- `CubeField As HRESULT  (read-only)`
- `CurrentPageName As HRESULT  (read/write)`
- `StandardFormula As HRESULT  (read/write)`
- `HiddenItemsList As HRESULT  (read/write)`
- `DatabaseSort As HRESULT  (read/write)`
- `IsMemberProperty As HRESULT  (read-only)`
- `PropertyParentField As HRESULT  (read-only)`
- `PropertyOrder As HRESULT  (read/write)`
- `EnableItemSelection As HRESULT  (read/write)`
- `CurrentPageList As HRESULT  (read/write)`
- `Hidden As HRESULT  (read/write)`
- `UseMemberPropertyAsCaption As HRESULT  (read/write)`
- `MemberPropertyCaption As HRESULT  (read/write)`
- `DisplayAsTooltip As HRESULT  (read/write)`
- `DisplayInReport As HRESULT  (read/write)`
- `DisplayAsCaption As HRESULT  (read-only)`
- `LayoutCompactRow As HRESULT  (read/write)`
- `IncludeNewItemsInFilter As HRESULT  (read/write)`
- `VisibleItemsList As HRESULT  (read/write)`
- `PivotFilters As HRESULT  (read-only)`
- `AutoSortPivotLine As HRESULT  (read-only)`
- `AutoSortCustomSubtotal As HRESULT  (read-only)`
- `ShowingInAxis As HRESULT  (read-only)`
- `EnableMultiplePageItems As HRESULT  (read/write)`
- `AllItemsVisible As HRESULT  (read-only)`
- `SourceCaption As HRESULT  (read-only)`
- `ShowDetail As HRESULT  (read/write)`
- `RepeatLabels As HRESULT  (read/write)`

## Methods (12)

- `PivotItems([Index As Variant], RHS As Variant)`
- `CalculatedItems(RHS As CalculatedItems)`
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
