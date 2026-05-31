# SharedWorkspaceTasks

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C037A-0000-0000-C000-000000000046}  

A collection of the SharedWorkspaceTask objects in the current shared workspace site.

**Example:**

```vba
Dim swsTasks As Office.SharedWorkspaceTasks
    Set swsTasks = ActiveWorkbook.SharedWorkspace.Tasks
    MsgBox "There are " & swsTasks.Count & _
        " task(s) in the current shared workspace.", _
        vbInformation + vbOKOnly, _
        "Collection Information"
    Set swsTasks = Nothing
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SharedWorkspaceTasks object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SharedWorkspaceTasks object was created. Read-only.
- `Item As SharedWorkspaceTask  (read-only)`  
  Gets a SharedWorkspaceTask object from the Tasks collection of the shared workspace. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the SharedWorkspaceTasks object. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SharedWorkspaceTasks object. Read-only.
- `ItemCountExceeded As Boolean  (read-only)`  
  Gets a Boolean value that indicates whether the number of SharedWorkspaceTasks items in the collection has exceeded the 99 that can be displayed in the Shared Workspace task pane. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Add(Title As String, [Status As Variant], [Priority As Variant], [Assignee As Variant], [Description As Variant], [DueDate As Variant]) As SharedWorkspaceTask`  
  Adds a task to the list of tasks in a shared workspace. Returns a SharedWorkspaceTask object.
    - `Title As String` (required): The title of the new task.
    - `Status As Variant` (optional): The status of the new task. Default is msoSharedWorkspaceTaskNotStarted.
    - `Priority As Variant` (optional): The priority of the new task. Default is msoSharedWorkspaceTaskNormal.
    - `Assignee As Variant` (optional): The member to whom the new task is assigned.
    - `Description As Variant` (optional): The description of the new task.
    - `DueDate As Variant` (optional): The due date of the new task.
