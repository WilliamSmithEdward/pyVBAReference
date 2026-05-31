# Sync

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0386-0000-0000-C000-000000000046}  

The Sync property of the Document object in Microsoft Word, the Workbook object in Microsoft Excel, and the Presentation object in Microsoft PowerPoint returns a Sync object.

**Remarks:** Use the Sync object to manage the synchronization of the local and server copies of a shared document stored on a SharePoint site. The Status property returns important information about the current state of synchronization. Use the GetUpdate method to refresh the sync status. Use the LastSyncTime, ErrorType, and WorkspaceLastChangedBy properties to return additional information. See the Status property for additional information about the differences and conflicts that can exist between the local and server copies of shared documents. Use the PutUpdate method to save local changes to the server. Close and re-open the document to retrieve the latest version from the server when no local changes have been made. Use the ResolveConflict method to resolve differences between the local and the server copies, or the OpenVersion method to open a different version alongside the currently open local version of the document. The GetUpdate, PutUpdate, and ResolveConflict methods of the Sync object don't return status codes because they complete their tasks asynchronously.

**Example:**

```vba
Dim objSync As Office.Sync
    Dim strStatus As String
    Set objSync = ActiveDocument.Sync
    If objSync.Status > msoSyncStatusNoSharedWorkspace Then
        Select Case objSync.Status
            Case msoSyncStatusConflict
                objSync.ResolveConflict msoSyncConflictMerge
                ActiveDocument.Save
                objSync.ResolveConflict msoSyncConflictClientWins
                strStatus = "Conflict resolved by merging changes."
            Case msoSyncStatusError
                strStatus = "Last error type: " & objSync.ErrorType
            Case msoSyncStatusLatest
                strStatus = "Document copies already in sync."
            Case msoSyncStatusLocalChanges
                objSync.PutUpdate
                strStatus = "Local changes saved to server."
            Case msoSyncStatusNewerAvailable
                objSync.GetUpdate
                strStatus = "Local copy updated from server."
            Case msoSyncStatusSuspended
                objSync.Unsuspend
                strStatus = "Synchronization resumed."
        End Select
    Else
        strStatus = "Not a shared workspace document."
    End If
    MsgBox strStatus, vbInformation + vbOKOnly, "Sync Information"
    Set objSync = Nothing
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the Sync object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the Sync object was created. Read-only.
- `Status As MsoSyncStatusType  (read-only)`  
  Gets the status of the synchronization of the local copy of the active document with the server copy. Read-only.
- `WorkspaceLastChangedBy As String  (read-only)`  
  Displays the display name of the user who last saved changes to the server copy of a shared document. Read-only.
- `LastSyncTime As Variant  (read-only)`  
  Gets the date and time when the local copy of the active document was last synchronized with the server copy. Read-only.
- `ErrorType As MsoSyncErrorType  (read-only)`  
  Gets an MsoSyncErrorType constant that indicates the type of the most recent document synchronization error. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the Sync object. Read-only.

## Methods (5)

- `GetUpdate()`  
  Compares the local version of the shared document to the version on the server.
- `PutUpdate()`  
  Updates the server copy of the shared document with the local copy.
- `OpenVersion(SyncVersionType As MsoSyncVersionType)`  
  Opens a different version of the shared document alongside the currently open local version.
    - `SyncVersionType As MsoSyncVersionType` (required): Represents the type of version.
- `ResolveConflict(SyncConflictResolution As MsoSyncConflictResolutionType)`  
  Resolves conflicts between the local and the server copies of a shared document.
    - `SyncConflictResolution As MsoSyncConflictResolutionType` (required): Specifies how conflicts should be resolved.
- `Unsuspend()`  
  Resumes synchronization between the local copy and the server copy of a shared document.
