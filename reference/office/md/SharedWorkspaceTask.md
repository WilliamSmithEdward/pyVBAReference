# SharedWorkspaceTask

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0379-0000-0000-C000-000000000046}  

The SharedWorkspaceTask object represents a task in a shared document workspace. Member of the SharedWorkspaceTasks collection.

**Remarks:** Use the SharedWorkspaceTask object to manage tasks assigned to the members who are collaborating on the documents in the shared workspace. Use the Item (_index_) property of the SharedWorkspaceTasks collection to return a specific SharedWorkspaceTask object. Use the Title property to set the text of the task that appears on the Tasks tab of the Shared Workspace task pane and on the shared workspace webpage. Use the Description property to supply additional information about the task. Assign the task to a member of the workspace by using the AssignedTo property and the member's domain user name. Specify a due date for the task by using the DueDate property. Use the enumerations for task Priority and Status to indicate the relative importance of the task and to update the task's status. Use the Save method to upload changes to the server after you modify properties of the SharedWorkspaceTask object. Use the CreatedBy, CreatedDate, ModifiedBy, and ModifiedDate properties to return information about the history of each task.

**Example:**

```vba
Dim swsTask As Office.SharedWorkspaceTask
    Dim strTaskInfo As String
    strTaskInfo = "The shared workspace contains " & _
    ActiveWorkbook.SharedWorkspace.Tasks.Count & " Task(s)." & vbCrLf
    For Each swsTask In ActiveWorkbook.SharedWorkspace.Tasks
        strTaskInfo = strTaskInfo & swsTask.Title & vbCrLf & _
            " - Description: " & swsTask.Description & vbCrLf & _
            " - Assigned to: " & swsTask.AssignedTo & vbCrLf & _
            " - Due date: " & swsTask.DueDate & vbCrLf & _
            " - Priority: " & swsTask.Priority & vbCrLf & _
            " - Status: " & swsTask.Status & vbCrLf
    Next
    MsgBox strTaskInfo, vbInformation + vbOKOnly, _
        "Tasks in Shared Workspace"
    Set swsTask = Nothing
```

## Properties (13)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SharedWorkspaceTask object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SharedWorkspaceTask object was created. Read-only.
- `Title As String  (read/write)`  
  Sets or gets the title of a SharedWorkspaceTask object. Read/write.
- `AssignedTo As String  (read/write)`  
  Gets or sets a value containing the name of the one assigned to the task represented by the SharedWorkspaceTask object. Read/write.
- `Status As MsoSharedWorkspaceTaskStatus  (read/write)`  
  Gets or sets the status of the specified shared workspace task. Read/write.
- `Priority As MsoSharedWorkspaceTaskPriority  (read/write)`  
  Gets or sets the status of the specified shared workspace task. Read/write.
- `Description As String  (read/write)`  
  Gets or sets a descriptive String value for the specified SharedWorkspaceLink or SharedWorkspaceTask object. Read/write.
- `DueDate As Variant  (read/write)`  
  Gets or sets the optional due date and time of a SharedWorkspaceTask object. Read/write.
- `CreatedBy As String  (read-only)`  
  Gets the display name of the member who created the shared workspace object. Read-only.
- `CreatedDate As Variant  (read-only)`  
  Gets the date and time when the shared workspace object was created. Read-only.
- `ModifiedBy As String  (read-only)`  
  Gets the name of the user who last modified the object. Read-only.
- `ModifiedDate As Variant  (read-only)`  
  Gets the date and time when the SharedWorkspaceTask object was last modified. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SharedWorkspaceTask object. Read-only.

## Methods (2)

- `Save()`  
  Uploads changes made programmatically to a shared server.
- `Delete()`  
  Deletes the current SharedWorkspaceTask object.
