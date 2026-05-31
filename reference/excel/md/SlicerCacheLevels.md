# SlicerCacheLevels

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244C5-0000-0000-C000-000000000046}  

Represents the collection of hierarchy levels for the OLAP data source that is filtered by a slicer.

**Remarks:** When a slicer is used to filter an OLAP data source, its parent slicer cache can contain multiple hierarchy levels from the data source. Use the SlicerCacheLevels collection of the parent SlicerCache object to access the SlicerCacheLevel objects that represent these hierarchy levels. This collection is not accessible for non-OLAP data sources.

**Example:**

```vba
ActiveWorkbook.SlicerCaches("Slicer_Customer_Geography"). _
 SlicerCacheLevels("[Customer].[Customer Geography].[Country]")
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the SlicerCache object that is the parent of the specified SlicerCacheLevels object. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of levels in the specified SlicerCacheLevels collection, which represents the number of levels in the associated SlicerCache object.
- `Item As SlicerCacheLevel  (read-only)`  
  Returns the specified SlicerCacheLevel object from the collection, or if no level is specified, returns the first SlicerCacheLevel object in the collection.
- `_Default As SlicerCacheLevel  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
