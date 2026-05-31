# MsoSyncConflictResolutionType

**Type:** Enumeration  
**Library:** Microsoft Office 16.0 Object Library  

Specifies how conflicts should be resolved when synchronizing a shared document. Used with the ResolveConflict method of the Sync object.

## Constants (3)

- `msoSyncConflictClientWins` = 0  
  Replace the server copy with the local copy.
- `msoSyncConflictServerWins` = 1  
  Replace the local copy with the server copy.
- `msoSyncConflictMerge` = 2  
  Merge changes made to the server copy into the local copy. To resolve the conflict with the merged changes winning, you must save the active document after merging changes, and then call the ResolveConflict method again with the msoSyncConflictClientWins option.
