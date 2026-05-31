# WorkflowTask

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD900-0000-0000-C000-000000000046}  

Represents a single workflow task in a WorkflowTasks collection.

**Example:**

```vba
Sub DisplayWorkTask()
Dim objWorkflowTasks As WorkflowTasks
Dim objWorkflowTask As WorkflowTask
Dim cnt As Integer

Set objWorkflowTasks = Document.GetWorkflowTasks()

For cnt = 1 To objWorkflowTasks.Count
 Debug.Print objWorkflowTask(cnt).Name
Next

Set objWorkflowTask = objWorkflowTasks(1)
objWorkflowTask.Show

End Sub
```

## Properties (11)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the WorkflowTask object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the WorkflowTask object was created. Read-only.
- `Id As String  (read-only)`  
  Gets the ID of the Sharepoint list item. Read-only.
- `ListID As String  (read-only)`  
  Gets the ID of the list containing the workflow task. Read-only.
- `WorkflowID As String  (read-only)`  
  Gets the ID of the workflow associated with a workflow task. Read-only.
- `Name As String  (read-only)`  
  Gets the name of the WorkflowTask object. Read-only.
- `Description As String  (read-only)`  
  Gets the description of a workflow task. Read-only.
- `AssignedTo As String  (read-only)`  
  Gets the name of the person that the workflow task is assigned to. Read-only.
- `CreatedBy As String  (read-only)`  
  Gets the name of the person that created the workflow task. Read-only.
- `DueDate As Date  (read-only)`  
  Gets the date that a workflow task is due. Read-only.
- `CreatedDate As Date  (read-only)`  
  Gets the date that a workflow task was created. Read-only.

## Methods (1)

- `Show() As Long`  
  Displays a workflow task edit user interface for the specified WorkflowTask object.
