# SlicerItems

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244CA-0000-0000-C000-000000000046}  

Represents the collection of SlicerItem objects contained in a SlicerCache or SlicerCacheLevel object.

**Remarks:** To access the SlicerItems collection that represents the items in a slicer based on data in the workbook or non-OLAP external data, use the SlicerItems property of the SlicerCache object that is associated with the slicer. To access the SlicerItems collection that represents the items in a slicer based on an OLAP data connection, use the SlicerItems property of the SlicerCacheLevel object that represents a level of the hierarchy.

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent SlicerCache object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only.
- `Item As SlicerItem  (read-only)`  
  Returns a SlicerItem object from a collection. Read-only.
- `_Default As SlicerItem  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
