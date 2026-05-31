# DocumentInspector

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0393-0000-0000-C000-000000000046}  

Represents a Document Inspector module in a DocumentInspectors collection.

**Remarks:** The DocumentInspector object provides access to the Inspect and Fix methods. The DocumentInspector object represents custom Document Inspector modules and some "built-in" options. To see the list of built-in options, open the Document Inspector dialog box from the File menu. The first two options (Comments, Revisions, Versions, and Annotations and Document Properties and Personal Information) are not listed in the DocumentInspectors collection; instead, their functionality is available through the RemoveDocumentInformation method. The remaining options in the dialog box and any installed custom modules developed by third-party sources are available from the DocumentInspectors collection by specifying an index value.

**Example:**

```vba
Public Sub DI_InspectDocument()
Dim docStatus As MsoDocInspectorStatus
Dim result As String
ActiveDocument.DocumentInspectors(3).Inspect docStatus, results

MsgBox ("The inspection returned the following status " & docStatus & _
" with this result " & result)
End Sub
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the creator of the DocumentInspector object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the DocumentInspector object was created. Read-only.
- `Name As String  (read-only)`  
  Gets the name of the module represented by a DocumentInspector object. Read-only.
- `Description As String  (read-only)`  
  Gets the description of the DocumentInspector object. Read-only.
- `Parent As Object  (read-only)`  
  Gets an object that represents the parent of the DocumentInspector object. Read-only.

## Methods (2)

- `Inspect(Status As MsoDocInspectorStatus, Results As String)`  
  Inspects a document for specific information or document properties.
    - `Status As MsoDocInspectorStatus` (required): An enumeration representing the status of the document. _Status_ is an output parameter, which means that its value is returned when the method has completed its purpose.
    - `Results As String` (required): Contains a list of the information items or document properties found in the document.
- `Fix(Status As MsoDocInspectorStatus, Results As String)`  
  Performs an action on specific information items or document properties depending on the Document Inspector module specified.
    - `Status As MsoDocInspectorStatus` (required): An enumeration representing the status of the document. _Status_ is an output parameter, which means that its value is returned when the method has completed its purpose.
    - `Results As String` (required): Contains the results of the action. _Results_ is an output parameter.
