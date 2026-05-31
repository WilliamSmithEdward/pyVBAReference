# Comments

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024426-0000-0000-C000-000000000046}  

A collection of cell comments.

**Remarks:** Each comment is represented by a Comment object.

**Example:**

```vba
Set cmt = Worksheets(1).Comments
For Each c In cmt
 c.Visible = False
Next
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
- `_Default As Comment  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As Long) As Comment`  
  Returns a single object from a collection.
    - `Index As Long` (required): The index number for the object.
