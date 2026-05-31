# DocumentInspectors

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0392-0000-0000-C000-000000000046}  

Represents a collection of DocumentInspector objects.

**Remarks:** The DocumentInspectors collection is part of the Document object in Microsoft Word, the Workbook object in Excel, and the Presentation object in PowerPoint. A DocumentInspectors collection contains multiple DocumentInspector objects, one for some built-in options and each installed custom Document Inspector module.

**Example:**

```vba
Public Sub FixDocument()
Dim docStatus As MsoDocInspectorStatus
Dim results As String
 ActiveDocument.DocumentInspectors(3).Fix docStatus, results

 MsgBox docStatus
 MsgBox("The following items were removed " & results)

End Sub
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the creator of the DocumentInspectors object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the DocumentInspectors object was created. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Item As DocumentInspector  (read-only)`  
  Gets the DocumentInspector object specified by the index. Read-only.
- `Count As Long  (read-only)`  
  Gets the number of items in the DocumentInspectors object. Read-only.
- `Parent As Object  (read-only)`  
  Gets an object that represents the parent of a DocumentInspectors object. Read-only.
