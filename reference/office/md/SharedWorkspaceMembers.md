# SharedWorkspaceMembers

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0382-0000-0000-C000-000000000046}  

A collection of the SharedWorkspaceMember objects in the current shared workspace site.

**Example:**

```vba
Dim swsMembers As Office.SharedWorkspaceMembers
    Set swsMembers = ActiveWorkbook.SharedWorkspace.Members
    MsgBox "There are " & swsMembers.Count & _
        " member(s) in the current shared workspace.", _
        vbInformation + vbOKOnly, _
        "Collection Information"
    Set swsMembers = Nothing
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SharedWorkspaceMembers object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SharedWorkspaceMembers object was created. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Item As SharedWorkspaceMember  (read-only)`  
  Gets a SharedWorkspaceMember object from the Members collection of the shared workspace. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the SharedWorkspaceMembers object. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SharedWorkspaceMembers object. Read-only.
- `ItemCountExceeded As Boolean  (read-only)`  
  Gets a Boolean value that indicates whether the number of SharedWorkspaceMembers items in the collection has exceeded the 99 that can be displayed in the Shared Workspace task pane. Read-only.

## Methods (1)

- `Add(Email As String, DomainName As String, DisplayName As String, [Role As Variant]) As SharedWorkspaceMember`  
  Adds a member to the list of members in a shared workspace site. Returns a SharedWorkspaceMember object.
    - `Email As String` (required): The new member's email address in the format user@domain.com. Raises an error if the user is not a valid candidate for membership in the shared workspace site.
    - `DomainName As String` (required): The new member's Windows user name in the format domain\user.
    - `DisplayName As String` (required): The display name to display for the new member.
    - `Role As Variant` (optional): An optional role that determines the tasks that the new member can accomplish in the shared workspace site; for example, "Contributor." An invalid role name raises an error.
