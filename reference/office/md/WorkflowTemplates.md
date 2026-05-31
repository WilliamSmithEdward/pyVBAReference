# WorkflowTemplates

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD903-0000-0000-C000-000000000046}  

Represents a collection of WorkflowTemplate objects.

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

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the WorkflowTemplates object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the WorkflowTemplates object was created. Read-only.
- `Item As WorkflowTemplate  (read-only)`  
  Gets a WorkflowTemplate object from the WorkflowTemplates collection. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the WorkflowTemplates collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
