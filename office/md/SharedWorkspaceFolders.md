# SharedWorkspaceFolders

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C037E-0000-0000-C000-000000000046}  

A collection of the SharedWorkspaceFolder objects in the current shared workspace.

**Example:**

```vba
Dim swsFolders As Office.SharedWorkspaceFolders
    Set swsFolders = ActiveWorkbook.SharedWorkspace.Folders
    MsgBox "There are " & swsFolders.Count & _
        " folder(s) in the current shared workspace.", _
        vbInformation + vbOKOnly, _
        "Collection Information"
    Set swsFolders = Nothing
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SharedWorkspaceFile object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SharedWorkspaceFolders object was created. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Item As SharedWorkspaceFolder  (read-only)`  
  Gets a SharedWorkspaceFolder object from the Folders collection of the shared workspace. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the SharedWorkspaceFolders collection. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SharedWorkspaceFolders object. Read-only.
- `ItemCountExceeded As Boolean  (read-only)`  
  Gets a Boolean value that indicates whether the number of SharedWorkspaceFolders items in the collection has exceeded the 99 that can be displayed in the Shared Workspace task pane. Read-only.

## Methods (1)

- `Add(FolderName As String, [ParentFolder As Variant]) As SharedWorkspaceFolder`  
  Adds a folder to the document library in a shared workspace. Returns a SharedWorkspaceFolder object.
    - `FolderName As String` (required): The name of the folder to be added to the current shared workspace.
    - `ParentFolder As Variant` (optional): The subfolder in which to place the new folder, if not the main document library folder within the shared workspace. Add the folder to the main document library folder by leaving this optional argument empty.
