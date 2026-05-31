# SlicerItem

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244C9-0000-0000-C000-000000000046}  

Represents an item in a slicer.

**Remarks:** To access the SlicerItem object that represents the currently selected button in the slicer, use the ActiveItem property of the Slicer object. To access the SlicerItems collection that represents all the items in a slicer filtering a PivotTable, use the SlicerItems property of the SlicerCache object that is associated with the Slicer object. To access the SlicerItems collection that represents the items in a slicer filtering a level of an OLAP hierarchy, use the SlicerItems property of the SlicerCacheLevel object that represents that level of the hierarchy.

## Properties (10)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As SlicerCache  (read-only)`  
  Returns the parent SlicerCache object for the slicer item. Read-only.
- `Caption As String  (read-only)`  
  Returns the label text for the slicer item. Read-only.
- `Name As String  (read-only)`  
  Returns the name of the slicer item. Read-only.
- `SourceName As Variant  (read-only)`  
  Returns the source name of the slicer item. Read-only.
- `SourceNameStandard As String  (read-only)`  
  Returns the slicer item's source name in standard English (United States) format settings. Read-only.
- `Value As String  (read-only)`  
  Returns the value of the slicer item. Read-only.
- `Selected As Boolean  (read/write)`  
  Returns or sets whether the slicer item is selected. Read/write for slicers connected to non-OLAP data sources. Read-only for slicers connected to OLAP data sources.
- `HasData As Boolean  (read-only)`  
  Returns whether the slicer item contains data that matches the current manual filter state. Read-only.
