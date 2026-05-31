# SharedWorkspaceFiles

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C037C-0000-0000-C000-000000000046}  

A collection of the SharedWorkspaceFile objects in the current shared workspace.

**Example:**

```vba
Dim swsFiles As Office.SharedWorkspaceFiles
    Set swsFiles = ActiveWorkbook.SharedWorkspace.Files
    MsgBox "There are " & swsFiles.Count & _
        " file(s)
        vbInformation + vbOKOnly, _
        "Collection Information"
    Set swsFiles = Nothing
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SharedWorkspaceFile object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SharedWorkspaceFiles object was created. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Item As SharedWorkspaceFile  (read-only)`  
  Gets a SharedWorkspaceFile object from the Files collection of the shared workspace. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the SharedWorkspaceFiles collection. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SharedWorkspaceFiles object. Read-only.
- `ItemCountExceeded As Boolean  (read-only)`  
  Gets a Boolean value indicating whether the number of files allowed in the shared workspace has been exceeded. Read-only.

## Methods (1)

- `Add(FileName As String, [ParentFolder As Variant], [OverwriteIfFileAlreadyExists As Variant], [KeepInSync As Variant]) As SharedWorkspaceFile`  
  Adds a file to the document library in a shared workspace. Returns a SharedWorkspaceFile object.
    - `FileName As String` (required): The path and file name of the file to be added to the current shared workspace.
    - `ParentFolder As Variant` (optional): The subfolder in which to place the file, if not the main document library folder within the shared workspace. Add the file to the main document library folder by leaving this optional argument empty.
    - `OverwriteIfFileAlreadyExists As Variant` (optional): True to overwrite an existing file by the same name. Default is False.
    - `KeepInSync As Variant` (optional): True to keep the local copy of the document synchronized with the copy in the shared workspace. Default is False.
