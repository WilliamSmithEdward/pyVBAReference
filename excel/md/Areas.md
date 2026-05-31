# Areas

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020860-0000-0000-C000-000000000046}  

A collection of the areas, or contiguous blocks of cells, within a selection.

**Remarks:** There's no singular Area object; individual members of the Areas collection are Range objects. The Areas collection contains one Range object for each discrete, contiguous range of cells within the selection. If the selection contains only one area, the Areas collection contains a single Range object that corresponds to that selection.

**Example:**

```vba
If Selection.Areas.Count <> 1 Then Selection.Clear
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As Range  (read-only)`  
  Returns a single Range object from the collection.
- `_NewEnum As IUnknown  (read-only)`
- `_Default As Range  (read-only)`
