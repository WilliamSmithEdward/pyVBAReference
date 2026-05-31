# UserPermission

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0375-0000-0000-C000-000000000046}  

Associates a set of permissions on the active document with a single user and an optional expiration date. Represents a member of the active document's Permission collection.

**Remarks:** Use the Add method of the Permission object to grant specific permissions on the active document to a new user, with an optional expiration date. Use the Remove method of the UserPermission object to remove a user and the user's permissions. While some permissions granted through the user interface (such as msoPermissionPrint) apply to all users, you can use the UserPermission object to assign them on a per-user basis with per-user expiration dates.

**Example:**

```vba
Dim irmPermission As Office.Permission
 Dim irmUserPerm As Office.UserPermission
 Dim strIRMInfo As String
 Set irmPermission = ActiveWorkbook.Permission
 If irmPermission.Enabled Then
 For Each irmUserPerm In irmPermission
 strIRMInfo = strIRMInfo & irmUserPerm.UserId & vbCrLf & _
 " - Permissions: " & irmUserPerm.Permission & vbCrLf & _
 " - Expiration Date: " & irmUserPerm.ExpirationDate & vbCrLf
 Next
 MsgBox strIRMInfo, _
 vbInformation + vbOKOnly, "IRM Information"
 Else
 MsgBox "This document is not restricted.", _
 vbInformation + vbOKOnly, "IRM Information"
 End If
 Set irmUserPerm = Nothing
 Set irmPermission = Nothing
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the UserPermission object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the UserPermission object was created. Read-only.
- `UserId As String  (read-only)`  
  Gets the email name of the user whose permissions on the active document are determined by the specified UserPermission object. Read-only.
- `Permission As Long  (read/write)`  
  Returns or sets an MsoPermission constant as a Long value representing the permissions on the active document assigned to the user associated with the specified UserPermission object. Read/write.
- `ExpirationDate As Variant  (read/write)`  
  Gets or sets the optional expiration date of the permissions on the active document assigned to the user associated with the specified UserPermission object. Read/write.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the UserPermission object. Read-only.

## Methods (1)

- `Remove()`  
  Removes the specified UserPermission object from the Permission collection of the active document.
