# UserAccessList

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002446C-0000-0000-C000-000000000046}  

A collection of UserAccess objects that represents the user access for protected ranges.

**Example:**

```vba
Sub UseDeleteAll()

 Dim wksSheet As Worksheet

 Set wksSheet = Application.ActiveSheet

 ' Notify the user of the number of users that can access the protected range.
 MsgBox wksSheet.Protection.AllowEditRanges(1).Users.Count

End Sub
```

## Properties (4)

- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As UserAccess  (read-only)`  
  Returns a single object from a collection.
- `_Default As UserAccess  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Add(Name As String, AllowEdit As Boolean) As UserAccess`  
  Adds a user access list.
    - `Name As String` (required): The name of the user access list.
    - `AllowEdit As Boolean` (required): True allows users on the access list to edit the editable ranges on a protected worksheet.
- `DeleteAll()`  
  Removes all users who have access to a protected range on a worksheet.
