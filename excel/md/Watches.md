# Watches

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024456-0000-0000-C000-000000000046}  

A collection of all the Watch objects in a specified application.

**Example:**

```vba
Sub AddWatch()

 With Application
 .Range("A1").Formula = 1
 .Range("A2").Formula = 2
 .Range("A3").Formula = "=Sum(A1:A2)"
 .Range("A3").Select
 .Watches.Add Source:=ActiveCell
 End With

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
- `_Default As Watch  (read-only)`
- `Item As Watch  (read-only)`  
  Returns a single object from a collection.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Add(Source As Variant) As Watch`  
  Adds a range that is tracked when the worksheet is recalculated.
    - `Source As Variant` (required): The source for the range.
- `Delete()`  
  Deletes the object.
