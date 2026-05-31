# WorkflowTasks

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD901-0000-0000-C000-000000000046}  

Represents a collection of WorkflowTask objects.

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

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the WorkflowTasks object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the WorkflowTasks object was created. Read-only.
- `Item As WorkflowTask  (read-only)`  
  Gets a WorkflowTask object from the WorkflowTasks collection. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the WorkflowTasks collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
