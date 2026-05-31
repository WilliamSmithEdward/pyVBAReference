# AllowEditRange

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002446B-0000-0000-C000-000000000046}  

Represents the cells that can be edited on a protected worksheet.

**Remarks:** Use the Add method or the Item property of the AllowEditRanges collection to return an AllowEditRange object. After an AllowEditRange object has been returned, you can use the ChangePassword method to change the password to access a range that can be edited on a protected worksheet.

**Example:**

```vba
Sub UseChangePassword()

 Dim wksOne As Worksheet
 Dim wksPassword As String

 Set wksOne = Application.ActiveSheet

 wksPassword = InputBox ("Enter password for the worksheet")

 ' Establish a range that can allow edits
 ' on the protected worksheet.
 wksOne.Protection.AllowEditRanges.Add _
 Title:="Classified", _
 Range:=Range("A1:A4"), _
 Password:=wksPassword

 MsgBox "Cells A1 to A4 can be edited on the protected worksheet."

 ' Change the password.

 wksPassword = InputBox ("Enter the new password for the worksheet")

 wksOne.Protection.AllowEditRanges(1).ChangePassword _
 Password:=wksPassword

 MsgBox "The password for these cells has been changed."

End Sub
```

## Properties (3)

- `Title As String  (read/write)`  
  Returns or sets the title of the range of cells that can edited on a protected sheet. Read/write String.
- `Range As Range  (read/write)`  
  Returns a Range object that represents a subset of the ranges that can be edited on a protected worksheet.
- `Users As UserAccessList  (read-only)`  
  Returns a UserAccessList object for the protected range on a worksheet.

## Methods (3)

- `ChangePassword(Password As String)`  
  Changes the password for a range that can be edited on a protected worksheet.
    - `Password As String` (required): The new password.
- `Delete()`  
  Deletes the object.
- `Unprotect([Password As Variant])`  
  Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.
    - `Password As Variant` (optional): A string that denotes the case-sensitive password to use to unprotect the range of cells. If the range isn't protected with a password, this argument is ignored.
