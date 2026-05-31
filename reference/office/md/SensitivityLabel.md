# SensitivityLabel

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {CA42DA75-D621-4476-B8F9-9C4ADAE77085}  

Represents a wrapper object for accessing sensitivity label on the active document.

**Remarks:** SensitivityLabel applied on a document requires the use of policy defined by organization administrator. The organization is identified by using an identity of an Office Account signed into Office.

## Properties (3)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SensitivityLabel object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SensitivityLabel object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SensitivityLabel object. Read-only.

## Methods (3)

- `GetLabel() As LabelInfo`  
  Gets the current label information that exists on the document for the user.
- `SetLabel(LabelInfo As LabelInfo, Context As Object)`  
  Sets the label information on the document for the user.
    - `LabelInfo As LabelInfo` (required): The label information that needs to be set on the document.
    - `Context As Object` (required): A caller defined context that can be returned with LabelChanged event to help ensure that the event was raised because of the SetLabel call.
- `CreateLabelInfo() As LabelInfo`  
  Creates a new LabelInfo object that can be passed to SetLabel method.

## Events (1)

- `LabelChanged(OldLabelInfo As LabelInfo, NewLabelInfo As LabelInfo, HResult As Long, Context As Object)`  
  Raised when a Label is changed on the document.
    - `OldLabelInfo As LabelInfo` (required): Previous label information that existed on the document.
    - `NewLabelInfo As LabelInfo` (required): New label information that was applied on the document.
    - `HResult As Long` (required): An integer representing the error code.
    - `Context As Object` (required): The _context_ object that was set with SetLabel call.
