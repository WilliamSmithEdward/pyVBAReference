# SharedWorkspaceLinks

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0380-0000-0000-C000-000000000046}  

A collection of the SharedWorkspaceLink objects in the current shared workspace.

**Example:**

```vba
Dim swsLinks As Office.SharedWorkspaceLinks
    Set swsLinks = ActiveWorkbook.SharedWorkspace.Links
    MsgBox "There are " & swsLinks.Count & _
        " link(s) in the current shared workspace.", _
        vbInformation + vbOKOnly, _
        "Collection Information"
    Set swsLinks = Nothing
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SharedWorkspaceLinks object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SharedWorkspaceLinks object was created. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Item As SharedWorkspaceLink  (read-only)`  
  Gets a SharedWorkspaceLink object from the Links collection of the shared workspace. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the SharedWorkspaceLinks object. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SharedWorkspaceLinks object. Read-only.
- `ItemCountExceeded As Boolean  (read-only)`  
  Gets a Boolean value that indicates whether the number of SharedWorkspaceLinks items in the collection has exceeded the 99 that can be displayed in the Shared Workspace task pane. Read-only.

## Methods (1)

- `Add(URL As String, [Description As Variant], [Notes As Variant]) As SharedWorkspaceLink`  
  Adds a link to the list of links in a shared workspace.
    - `URL As String` (required): The address of the website to which a link is being added.
    - `Description As Variant` (optional): Description of the link.
    - `Notes As Variant` (optional): Notes about the link.
