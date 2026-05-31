# SharedWorkspaceMember

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0381-0000-0000-C000-000000000046}  

Represents a user who has rights in a shared document workspace site.

**Remarks:** Use the SharedWorkspaceMember object to manage users who have rights to participate in a shared workspace and to collaborate on the shared documents saved in the workspace site. The Role specified when the user is added as a member of the workspace (for example, "Reader" or "Contributor") determines that user's rights in the workspace and cannot be accessed or modified later through properties of the SharedWorkspaceMember object. Use the Item (_index_) property of the SharedWorkspaceMembers collection to return a specific SharedWorkspaceMember object. Use the SharedWorkspaceMember object's three distinct name properties to retrieve identifying information about the member. - The Name property returns the members display name. - The Email property returns the member's email address. - The DomainName property returns the member's domain and user name in the format domain\user.

**Example:**

```vba
Dim swsMember As Office.SharedWorkspaceMember
    Dim strMemberInfo As String
    strMemberInfo = "The shared workspace contains " & _
        ActiveWorkbook.SharedWorkspace.Members.Count & " member(s)." & vbCrLf
    If ActiveWorkbook.SharedWorkspace.Members.Count > 0 Then
        For Each swsMember In ActiveWorkbook.SharedWorkspace.Members
            strMemberInfo = strMemberInfo & swsMember.Name & vbCrLf & _
                " - " & swsMember.DomainName & vbCrLf & _
                " - " & swsMember.Email & vbCrLf
        Next
    End If
    MsgBox strMemberInfo, vbInformation + vbOKOnly, _
        "Members in Shared Workspace"
    Set swsMember = Nothing
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SharedWorkspaceMember object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SharedWorkspaceMember object was created. Read-only.
- `DomainName As String  (read-only)`  
  Gets the domain and user name of the specified SharedWorkspaceMember object in the format domain\user. Read-only.
- `Name As String  (read-only)`  
  Gets the display name of the shared workspace member. Read-only.
- `Email As String  (read-only)`  
  Gets the email name of the specified SharedWorkspaceMember in the format user@domain.com. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SharedWorkspaceMember object. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the current SharedWorkspaceMember object.
