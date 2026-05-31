# SlicerCacheLevel

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244C6-0000-0000-C000-000000000046}  

Represents a level of a hierarchy in an OLAP data source being filtered by a slicer.

**Remarks:** When a slicer is used to filter an OLAP data source, the slicer cache that it is associated with can contain multiple hierarchy levels from the data source. Use the SlicerCacheLevels collection of the parent SlicerCache object to access the SlicerCacheLevel objects that represent these hierarchy levels.

## Properties (10)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the SlicerCache object for the specified SlicerCacheLevel object. Read-only.
- `SlicerItems As SlicerItems  (read-only)`  
  Returns a SlicerItems collection that contains the collection of all slicer items for the specified level. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of SlicerItem objects in the specified SlicerCacheLevel object. Read-only.
- `Ordinal As Long  (read-only)`  
  Returns the one-based ordinal value of the hierarchy level represented by the specified SlicerCacheLevel object. Read-only.
- `Name As String  (read-only)`  
  Returns the MDX unique name of the OLAP hierarchy level represented by the object.
- `CrossFilterType As XlSlicerCrossFilterType  (read/write)`  
  Returns or sets whether a slicer is participating in cross filtering with other slicers that share the same slicer cache, and how cross filtering is displayed. Read/write.
- `SortItems As XlSlicerSort  (read/write)`  
  Returns or sets the sort order of the items in the slicer. Read/write.
- `VisibleSlicerItemsList As Variant  (read-only)`  
  Returns the list of slicer items that are currently included in the slicer filter. Read-only.
