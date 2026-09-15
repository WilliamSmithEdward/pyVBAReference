# UndoRecord

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {E598E358-2852-42D4-8775-160BD91B7244}  

Provides an entry point into the undo stack.

**Remarks:** Use the UndoRecord object to create and modify custom undo records in the Word undo stack.

**Example:**

```vba
Dim objUndo As UndoRecord
Set objUndo = Application.UndoRecord
```

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified UndoRecord object.
- `IsRecordingCustomRecord As Boolean  (read-only)`  
  Returns a Boolean that specifies whether a custom undo action is being recorded. Read-only.
- `CustomRecordName As String  (read-only)`  
  Returns a String that specifies the entry that appears on the undo stack when all custom undo actions have completed. Read-only.
- `CustomRecordLevel As Long  (read-only)`  
  Returns a Long that specifies the number of custom undo action calls that are currently active. Read-only.

## Methods (2)

- `StartCustomRecord([Name As String])`  
  Initiates the creation of a custom undo record.
    - `Name As String` (optional): Specifies the name of the custom undo record. This string is limited to 64 characters. If a longer string is supplied, the string is truncated to 64 characters.
- `EndCustomRecord()`  
  Completes the creation of a custom undo record.
