# UserAccess

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002446D-0000-0000-C000-000000000046}  

Represents the user access for a protected range.

**Example:**

```vba
Sub UseAllowEditRanges()

 Dim wksSheet As Worksheet

 Set wksSheet = Application.ActiveSheet

 ' Add a range that can be edited on the protected worksheet.
 wksSheet.Protection.AllowEditRanges.Add "Test", Range("A1")

 ' Notify the user the title of the range that can be edited.
 MsgBox wksSheet.Protection.AllowEditRanges(1).Title

End Sub
```

## Properties (2)

- `Name As String  (read-only)`  
  Returns or sets a String value that represents the name of the object.
- `AllowEdit As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if the user is allowed access to the specified range on a protected worksheet.

## Methods (1)

- `Delete()`  
  Deletes the object.
