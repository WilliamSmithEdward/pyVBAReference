# Conflict

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {6215E4B1-545A-406E-9824-0A5B5AC8AD21}  

Represents a conflicting edit in a co authored document. The type of a Conflict object is specified by the WdRevisionType enumeration.

**Remarks:** Although co authoring in Word is designed to minimize conflicts, conflicts can sometimes occur when editing a document that has co authoring enabled. A conflict occurs when Word requires user input to resolve a merge. For example, conflicts could potentially occur when a user opens a co authored document from the server, works offline, and once online again, saves the document back to the server. As another example, conflicts can sometimes occur when more than one person works on the same document range at exactly the same time.

**Example:**

```vba
Dim con as Conflict

For Each con in ActiveDocument.CoAuthoring.Conflicts
MsgBox con.Type
Next con
```

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Conflict object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained in the specified object. Read-only.
- `Type As WdRevisionType  (read-only)`  
  Returns the WdRevisionTypefor the Conflict object. Read-only.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.

## Methods (2)

- `Accept()`  
  Accepts the user specified conflict change, and removes the conflict.
- `Reject()`  
  Rejects the user change, removes the conflict, and accepts the server copy of the change for the conflict.
