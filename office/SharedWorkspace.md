# SharedWorkspace

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0385-0000-0000-C000-000000000046}  

## Properties (13)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Name As String  (read/write)`
- `Members As SharedWorkspaceMembers  (read-only)`
- `Tasks As SharedWorkspaceTasks  (read-only)`
- `Files As SharedWorkspaceFiles  (read-only)`
- `Folders As SharedWorkspaceFolders  (read-only)`
- `Links As SharedWorkspaceLinks  (read-only)`
- `Parent As Object  (read-only)`
- `URL As String  (read-only)`
- `Connected As Boolean  (read-only)`
- `LastRefreshed As Variant  (read-only)`
- `SourceURL As String  (read/write)`

## Methods (5)

- `Refresh()`
- `CreateNew([URL As Variant], [Name As Variant])`
- `Delete()`
- `RemoveDocument()`
- `Disconnect()`
