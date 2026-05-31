# Panes

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020894-0000-0000-C000-000000000046}  

A collection of all the Pane objects shown in the specified window.

**Remarks:** Pane objects exist only for worksheets and Microsoft Excel 4.0 macro sheets.

**Example:**

```vba
If ActiveWindow.Panes.Count > 1 Then _
 ActiveWindow.FreezePanes = True
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
- `Item As Pane  (read-only)`  
  Returns a single object from a collection.
- `_Default As Pane  (read-only)`
