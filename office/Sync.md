# Sync

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0386-0000-0000-C000-000000000046}  

## Properties (7)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Status As MsoSyncStatusType  (read-only)`
- `WorkspaceLastChangedBy As String  (read-only)`
- `LastSyncTime As Variant  (read-only)`
- `ErrorType As MsoSyncErrorType  (read-only)`
- `Parent As Object  (read-only)`

## Methods (5)

- `GetUpdate()`
- `PutUpdate()`
- `OpenVersion(SyncVersionType As MsoSyncVersionType)`
- `ResolveConflict(SyncConflictResolution As MsoSyncConflictResolutionType)`
- `Unsuspend()`
