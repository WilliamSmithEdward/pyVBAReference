# MsoSyncStatusType

**Type:** Enumeration  
**Library:** Microsoft Office 16.0 Object Library  

Specifies the status of the synchronization of the local copy of the active document with the server copy. Used with the Status property of the Sync object.

## Constants (8)

- `msoSyncStatusNoSharedWorkspace` = 0  
  No shared workspace.
- `msoSyncStatusNotRoaming` = 0  
  No synchronization is needed.
- `msoSyncStatusLatest` = 1  
  Documents are already in sync.
- `msoSyncStatusNewerAvailable` = 2  
  Only server copy has changes.
- `msoSyncStatusLocalChanges` = 3  
  Only local copy has changes.
- `msoSyncStatusConflict` = 4  
  Both the local and the server copies have changes.
- `msoSyncStatusSuspended` = 5  
  Synchronization has been suspended.
- `msoSyncStatusError` = 6  
  An error occurred. Use the ErrorType property of the Sync object to determine the exact error.
