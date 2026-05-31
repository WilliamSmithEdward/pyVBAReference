# SlicerCache

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244C4-0000-0000-C000-000000000046}  

Represents the current filter state for a slicer, and information about which PivotCache or WorkbookConnection object the slicer is connected to.

**Remarks:** Use the SlicerCaches property of the Workbook object to access the SlicerCaches collection of SlicerCache objects in a workbook. Each slicer has a base SlicerCache object that represents the items displayed in the slicer and the current user interface state of the tiles displayed with their corresponding item captions. Each slicer control that the user sees in Excel is represented by a Slicer object that has a SlicerCache object associated with it.

**Example:**

```vba
With ActiveWorkbook
 .SlicerCaches.Add("AdventureWorks", _
 "[Customer].[Customer Geography]").Slicers.Add SlicerDestination:="Sheet2", _
 Level:="[Customer].[Customer Geography].[Country]", Caption:="Country"
End With
```

## Properties (25)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent SlicerCaches object for the specified SlicerCache object. Read-only.
- `Index As Long  (read-only)`  
  Returns the index of the specified SlicerCache object in the SlicerCaches collection. Read-only.
- `OLAP As Boolean  (read-only)`  
  Returns whether the slicer associated with the specified slicer cache is based on an OLAP data source. Read-only.
- `SourceType As XlPivotTableSourceType  (read-only)`  
  Returns the kind of data source that the slicer is connected to. Read-only.
- `WorkbookConnection As WorkbookConnection  (read-only)`  
  Gets or sets the WorkbookConnection object that represents the data connection used by the specified slicer. Read/write.
- `Slicers As Slicers  (read-only)`  
  Returns a Slicers collection that contains the collection of Slicer objects associated with the specified SlicerCache object. Read-only.
- `PivotTables As SlicerPivotTables  (read-only)`  
  Returns a SlicerPivotTables collection that contains information about the PivotTables that the slicer cache is currently filtering. Read-only.
- `SlicerCacheLevels As SlicerCacheLevels  (read-only)`  
  Returns the collection of SlicerCacheLevel objects that represent the levels of an OLAP hierarchy on which the specified slicer cache is based. Read-only.
- `Name As String  (read/write)`  
  Returns or sets the name of the slicer cache.
- `VisibleSlicerItems As SlicerItems  (read-only)`  
  Returns a SlicerItems collection that contains the collection of all the visible items in the specified slicer cache. Read-only.
- `VisibleSlicerItemsList As Variant  (read/write)`  
  Returns or sets the list of MDX unique names for members at all levels of the hierarchy where manual filtering is applied. Read/write.
- `SlicerItems As SlicerItems  (read-only)`  
  Returns a SlicerItems collection that contains the collection of all items in the slicer cache. Read-only.
- `CrossFilterType As XlSlicerCrossFilterType  (read/write)`  
  Returns or sets whether a slicer is participating in cross filtering with other slicers that share the same slicer cache, and how cross filtering is displayed. Read/write.
- `SortItems As XlSlicerSort  (read/write)`  
  Returns or sets the sort order of the items in the slicer. Read/write XlSlicerSort.
- `SourceName As String  (read-only)`  
  Returns the name of the data source that the slicer is connected to. Read-only.
- `SortUsingCustomLists As Boolean  (read/write)`  
  Returns or sets whether items in the specified slicer cache will be sorted by the custom lists. Read/write.
- `ShowAllItems As Boolean  (read/write)`  
  Returns or sets whether slicers connected to the specified slicer cache display items that have been deleted from the corresponding PivotCache. Read/write.
- `TimelineState As TimelineState  (read-only)`  
  The timeline-specific state of the SlicerCache object. Read-only.
- `SlicerCacheType As XlSlicerCacheType  (read-only)`  
  Returns the type of the slicer cache: slicer (xlSlicer) or timeline (xlTimeline) . Read-only XlSlicerCacheType.
- `FilterCleared As Boolean  (read-only)`  
  Returns whether the slicer or timeline filter state is cleared. Read-only Boolean.
- `List As Boolean  (read-only)`  
  True if the slicer cache is for a slicer on a table; otherwise, False. Read-only Boolean.
- `RequireManualUpdate As Boolean  (read/write)`  
  True when manual updates of the slicer cache are required. Read/write Boolean.
- `ListObject As ListObject  (read-only)`  
  Returns a ListObject object for the QueryTable object. Read-only.

## Methods (4)

- `ClearManualFilter()`  
  Clears the manual filter for the slicer cache.
- `Delete()`  
  Deletes the specified slicer cache and the slicers associated with it.
- `ClearAllFilters()`  
  Clears the filter for either slicer or timeline, depending on the slicer cache type.
- `ClearDateFilter()`  
  Clears the filter for a timeline (date filter).
