# Permission

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0376-0000-0000-C000-000000000046}  

The Permission property of the Document object in Microsoft Word, a Workbook object in Microsoft Excel, and a Presentation object in Microsoft PowerPoint returns a Permission object.

**Remarks:** Use the Permission object to restrict permissions to the active document and to return or set specific permissions settings. The Permission object gives access to a collection of UserPermission objects. Use the UserPermission object to associate specific sets of rights with individual users. While some permissions granted through the user interface (such as msoPermissionPrint) apply to all users, you can use the UserPermission object to assign them on a per-user basis with per-user expiration dates. Microsoft Office Information Rights Management supports the use of administrative permission policies that list users and groups and their document permissions. Use the ApplyPolicy method to apply a permission policy, and the PermissionFromPolicy, PolicyName, and PolicyDescription properties to return policy information. The Permission object model is available whether permissions are restricted on the active document or not. The Permission property of the Document, Workbook, and Presentation objects does not return Nothing when the active document does not have restricted permissions. Use the Enabled property to determine whether a document has restricted permissions.

**Example:**

```vba
Dim irmPermission As Office.Permission
 Dim strIRMInfo As String
 Set irmPermission = ActiveWorkbook.Permission
 If irmPermission.Enabled Then
 strIRMInfo = "Permissions are restricted on this document." & vbCrLf
 strIRMInfo = strIRMInfo & " View in trusted browser: " & _
 irmPermission.EnableTrustedBrowser & vbCrLf & _
 " Document author: " & irmPermission.DocumentAuthor & vbCrLf & _
 " Users with permissions: " & irmPermission.Count & vbCrLf & _
 " Cache licenses: " & irmPermission.StoreLicenses & vbCrLf & _
 " Request permission URL: " & irmPermission.RequestPermissionURL & vbCrLf
 If irmPermission.PermissionFromPolicy Then
 strIRMInfo = strIRMInfo & " Permissions applied from policy:" & vbCrLf & _
 " Policy name: " & irmPermission.PolicyName & vbCrLf & _
 " Policy description: " & irmPermission.PolicyDescription
 Else
 strIRMInfo = strIRMInfo & " Default permissions applied." & vbCrLf & _
 " Default policy name: " & irmPermission.PolicyName & vbCrLf & _
 " Default policy description: " & irmPermission.PolicyDescription
 End If
 Else
 strIRMInfo = "Permissions are NOT restricted on this document."
 End If
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information"
 Set irmPermission = Nothing
```

## Properties (16)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the Permission object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the Permission object was created. Read-only.
- `Item As UserPermission  (read-only)`  
  Gets a UserPermission object that is a member of the Permission collection. The UserPermission object associates a set of permissions on the active document with a single user and an optional expiration date. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the Permission object. Read-only.
- `EnableTrustedBrowser As Boolean  (read/write)`  
  Gets or sets a value indicating whether to enable a browser from a trusted source. Read/write.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the Permission object. Read-only.
- `Enabled As Boolean  (read/write)`  
  Gets or sets a Boolean value that indicates whether permissions are enabled on the active document. Read/write.
- `RequestPermissionURL As String  (read/write)`  
  Gets or sets the file or website URL to visit or the email address to contact for users who need additional permissions on the active document. Read/write.
- `PolicyName As String  (read-only)`  
  Gets the name of the permissions policy applied to the active document. Read-only.
- `PolicyDescription As String  (read-only)`  
  Gets the description of the permissions policy applied to the active document. Read-only.
- `StoreLicenses As Boolean  (read/write)`  
  Gets or sets a Boolean value that indicates whether the user's license to view the active document should be cached to allow offline viewing when the user cannot connect to a rights management server. Read/write.
- `DocumentAuthor As String  (read/write)`  
  Gets or sets the name in email form of the author of the active document. Read/write.
- `PermissionFromPolicy As Boolean  (read-only)`  
  Gets a Boolean value that indicates whether a permission policy has been applied to the active document. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `DoubleKeyEncryptionUrl As String  (read/write)`
- `SensitivityLabelId As String  (read/write)`  
  Gets or sets the sensitivity label id included in user defined protection from Microsoft Purview Information Protection. Read/write.

## Methods (3)

- `Add(UserId As String, [Permission As Variant], [ExpirationDate As Variant]) As UserPermission`  
  Creates a set of permissions on the active document for the specified user. Returns a UserPermission object.
    - `UserId As String` (required): The email address (in the format user@domain.com) of the user to whom permissions on the active document are being granted.
    - `Permission As Variant` (optional): The permissions on the active document that are being granted to the specified user.
    - `ExpirationDate As Variant` (optional): The expiration date for the permissions that are being granted. NOTE: this parameter is not used and will be ignored.
- `ApplyPolicy(FileName As String)`  
  Applies the specified permission policy to the active document.
    - `FileName As String` (required): The path and file name of the permission policy template file.
- `RemoveAll()`  
  Removes all UserPermission objects from the Permission collection of the active document.
