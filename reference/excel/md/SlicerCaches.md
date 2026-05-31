# SlicerCaches

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244C3-0000-0000-C000-000000000046}  

Represents the collection of slicer caches associated with the specified workbook.

**Remarks:** Use the Item property of the SlicerCaches collection to return a SlicerCache object associated with the specified Workbook object. A SlicerCache object can be retrieved by using either the value of the Index property or the Name property of the specified object.

**Example:**

```vba
ActiveWorkbook.SlicerCaches("Slicer_Country")
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent Workbook object for the collection. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection.
- `Item As SlicerCache  (read-only)`  
  Returns a single SlicerCache object from the collection.
- `_Default As SlicerCache  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Add2(Source As Variant, SourceField As Variant, [Name As Variant], [SlicerCacheType As Variant]) As SlicerCache`
