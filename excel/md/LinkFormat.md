# LinkFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024442-0000-0000-C000-000000000046}  

Contains linked OLE object properties.

**Remarks:** If the Shape object doesn't represent a linked object, the LinkFormat property of the Shape object fails.

**Example:**

```vba
Worksheets(1).Shapes(1).LinkFormat.Update
```

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `AutoUpdate As Boolean  (read/write)`  
  True if the LinkFormat object is updated automatically when the source changes. Read-only Boolean.
- `Locked As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if the object is locked.

## Methods (1)

- `Update()`  
  Updates the link.
