# Errors

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002445C-0000-0000-C000-000000000046}  

Represents the various spreadsheet errors for a range.

**Remarks:** Use the Errors property of the Range object to return an Errors object.

**Example:**

```vba
Sub ErrorValue()

 ' Place a number written as text in cell A1.
 Range("A1").Formula = "'1"

 If Range("A1").Errors.Item(xlNumberAsText).Value = True Then
 MsgBox "Cell A1 has a number as text."
 Else
 MsgBox "Cell A1 is a number."
 End If

End Sub
```

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As Error  (read-only)`  
  Returns a single member of the Error object.
- `_Default As Error  (read-only)`
