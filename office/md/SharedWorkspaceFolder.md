# SharedWorkspaceFolder

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C037D-0000-0000-C000-000000000046}  

Represents a folder in a shared document workspace.

**Remarks:** Use the SharedWorkspaceFolder object to manage subfolders within the main document library folder of a shared workspace. The Count property of the SharedWorkspaceFolders collection does not include the workspace's main folder, and returns 0 (zero) if no subfolders have been created. The SharedWorkspaceFolder object does not expose the CreatedBy, CreatedDate, ModifiedBy, and ModifiedDate properties available on the SharedWorkspaceFile, SharedWorkspaceLink, and SharedWorkspaceTask objects. Use the Item (_index_) property of the SharedWorkspaceFolders collection to return a specific SharedWorkspaceFolder object.

**Example:**

```vba
Dim swsFolder As SharedWorkspaceFolder
    Set swsFolder = ActiveWorkbook.SharedWorkspace.Folders(1)
    MsgBox swsFolder.FolderName, vbInformation + vbOKOnly, "Folder Name"
    Set swsFolder = Nothing
```

## Properties (4)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SharedWorkspaceFolder object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SharedWorkspaceFolder object was created. Read-only.
- `FolderName As String  (read-only)`  
  Gets the name of a subfolder within the main document library folder of a shared workspace. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SharedWorkspaceFolder object. Read-only.

## Methods (1)

- `Delete([DeleteEventIfFolderContainsFiles As Variant])`  
  Deletes the current shared workspace folder and all data within it.
