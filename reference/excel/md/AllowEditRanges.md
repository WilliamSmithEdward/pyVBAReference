# AllowEditRanges

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002446A-0000-0000-C000-000000000046}  

A collection of all the AllowEditRange objects that represent the cells that can be edited on a protected worksheet.

**Remarks:** Use the AllowEditRanges property of the Protection object to return an AllowEditRanges collection. After an AllowEditRanges collection has been returned, you can use the Add method to add a range that can be edited on a protected worksheet.

**Example:**

```vba
Sub UseAllowEditRanges()

 Dim wksOne As Worksheet
 Dim wksPassword As String

 Set wksOne = Application.ActiveSheet

 ' Unprotect worksheet.
 wksOne.Unprotect

 wksPassword = InputBox ("Enter password for the worksheet")

 ' Establish a range that can allow edits
 ' on the protected worksheet.
 wksOne.Protection.AllowEditRanges.Add _
 Title:="Classified", _
 Range:=Range("A1:A4"), _
 Password:=wksPassword

 ' Notify the user
 ' the title and address of the range.
 With wksOne.Protection.AllowEditRanges.Item(1)
 MsgBox "Title of range: " & .Title
 MsgBox "Address of range: " & .Range.Address
 End With

End Sub
```

## Properties (4)

- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As AllowEditRange  (read-only)`  
  Returns a single object from a collection.
- `_Default As AllowEditRange  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Add(Title As String, Range As Range, [Password As Variant]) As AllowEditRange`  
  Adds a range that can be edited on a protected worksheet. Returns an AllowEditRange object.
    - `Title As String` (required): The title of the range.
    - `Range As Range` (required): Range object. The range allowed to be edited.
    - `Password As Variant` (optional): The password for the range.
