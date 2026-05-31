# CustomViews

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024422-0000-0000-C000-000000000046}  

A collection of custom workbook views.

**Remarks:** Each view is represented by a CustomView object.

**Example:**

```vba
ActiveWorkbook.CustomViews.Add "Summary", True, True
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `_Default As CustomView  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Item(ViewName As Variant) As CustomView`  
  Returns a single object from a collection.
- `Add(ViewName As String, [PrintSettings As Variant], [RowColSettings As Variant]) As CustomView`  
  Creates a new custom view.
    - `ViewName As String` (required): The name of the new view.
    - `PrintSettings As Variant` (optional): True to include print settings in the custom view.
    - `RowColSettings As Variant` (optional): True to include settings for hidden rows and columns (including filter information) in the custom view.
