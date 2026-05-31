# IDocumentInspector

**Type:** Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD706-0000-0000-C000-000000000046}  

Represents the interface through which the methods of an IDocumentInspector object are accessed.

**Remarks:** IDocumentInspector provides access to the Fix, GetInfo, and Inspect methods. The IDocumentInspector object is used with custom Document Inspector modules. This is in contrast to modules that are shipped with Microsoft Office.

## Methods (3)

- `GetInfo(Name As String, Desc As String)`  
  Gets information about a custom Document Inspector module.
    - `Name As String` (required): Represents the name of the module.
    - `Desc As String` (required): Represents the description of the module.
- `Inspect(Doc As Object, Status As MsoDocInspectorStatus, Result As String, Action As String)`  
  Inspects a document for specific information items or document properties by using a custom Document Inspector module.
    - `Doc As Object` (required): An object representing the container document.
    - `Status As MsoDocInspectorStatus` (required): An enumeration that represents the results of the inspection.
    - `Result As String` (required): Contains a list of the information items or document properties found in the document.
    - `Action As String` (required): Indicates to the user what action to take based on the results of the inspection.
- `Fix(Doc As Object, hwnd As Long, Status As MsoDocInspectorStatus, Result As String)`  
  Performs some action on specific information items or document properties by using a custom Document Inspector module.
    - `Doc As Object` (required): An object representing the container object.
    - `hwnd As Long` (required): Unique identifier of the active document window.
    - `Status As MsoDocInspectorStatus` (required): An enumeration that indicates the status of the action.
    - `Result As String` (required): Contains the results of the action.
