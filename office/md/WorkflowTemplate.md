# WorkflowTemplate

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD902-0000-0000-C000-000000000046}  

Represents one of the workflows available for the current document.

**Remarks:** A WorkflowTemplate object corresponds to one of the options displayed in the Start New Workflow dialog box. On a webpage, the workflow templates are displayed as a list of options.

**Example:**

```vba
Sub DisplayWorkTemplates()
Dim objWorkflowTemplates As WorkflowTemplates
Dim objWorkflowTemplate As WorkflowTemplate
Dim cnt As Integer

Set objWorkflowTemplates = Document.GetWorkflowTemplates()

For cnt = 1 To objWorkflowTemplates.Count
 Debug.Print objWorkflowTemplate(cnt).Name
Next

Set objWorkflowTemplate = objWorkflowTemplates(1)
objWorkflowTemplate.Show

End Sub
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the WorkflowTemplate object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the WorkflowTemplate object was created. Read-only.
- `Id As String  (read-only)`  
  Gets the ID of a template used to create a workflow instance. Read-only.
- `Name As String  (read-only)`  
  Gets the name of the WorkflowTemplate object. Read-only.
- `Description As String  (read-only)`  
  Gets the description of a workflow template. Read-only.
- `DocumentLibraryName As String  (read-only)`  
  Gets the name of the document library associated with the workflow template. Read-only.
- `DocumentLibraryURL As String  (read-only)`  
  Gets the URL address of the document library where workflow templates are stored. Read-only.

## Methods (1)

- `Show() As Long`  
  Displays a workflow-specific configuration user interface for the specified WorkflowTemplate object.
