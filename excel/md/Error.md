# Error

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002445D-0000-0000-C000-000000000046}  

Represents a spreadsheet error for a range.

**Remarks:** This object works for ranges containing only one cell. Use the Item property of the Errors object to return an Error object. After an Error object is returned, you can use the Value property in conjunction with the Errors property of the Range object to check whether a particular error checking option is enabled.

**Example:**

```vba
Sub CheckEmptyCells()

 Dim rngFormula As Range
 Set rngFormula = Application.Range("A1")

 ' Place a formula referencing empty cells.
 Range("A1").Formula = "=A2+A3"
 Application.ErrorCheckingOptions.EmptyCellReferences = True

 ' Perform check to see if EmptyCellReferences check is on.
 If rngFormula.Errors.Item(xlEmptyCellReferences).Value = True Then
 MsgBox "The empty cell references error checking feature is enabled."
 Else
 MsgBox "The empty cell references error checking feature is not on."
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
- `Value As Boolean  (read-only)`  
  Returns a Boolean value that indicates if all the validation criteria are met (that is, if the range contains valid data).
- `Ignore As Boolean  (read/write)`  
  Allows the user to set or return the state of an error checking option for a range. False enables an error checking option for a range. True disables an error checking option for a range. Read/write Boolean.
