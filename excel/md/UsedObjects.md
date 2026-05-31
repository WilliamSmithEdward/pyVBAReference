# UsedObjects

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024451-0000-0000-C000-000000000046}  

Represents objects that have been allocated in a workbook.

**Example:**

```vba
Sub CountUsedObjects()

 MsgBox "The number of used objects in this application is: " & _
 Application.UsedObjects.Count

End Sub
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
- `_NewEnum As IUnknown  (read-only)`
- `_Default As Object  (read-only)`
- `Item As Object  (read-only)`  
  Returns a single object from a collection.
