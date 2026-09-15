# CoAuthUpdate

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {9E6B5EC5-E8E4-40AF-9540-6203F71E2823}  

Represents a range of text that has been updated by a co author.

**Remarks:** When a document that has co authoring enabled is edited by more than one author, changes to the document by one author are pushed to other authors' versions of the document by using updates. When a co author performs an explicit document save (by pressing CTRL + S, for example), changes made by other co authors are merged into the document as updates. The CoAuthUpdates collection contains all changes that were merged into the document, where each change is a single update represented by a CoAuthUpdate object. The contents of the CoAuthUpdates collection remains the same until a co author performs another explicit document save. When the co author saves the document again, if there are no new changes from other co authors that are merged into the document, the CoAuthUpdates collection retains the same updates that were merged at the previous explicit save. If there are new changes that are merged into the document, the CoAuthUpdates collection contains the new updates for the document. Use a CoAuthUpdate object to retrieve an individual update from the CoAuthUpdates collection.

**Example:**

```vba
Dim caUpdate As CoAuthUpdate
Dim strText As String

For Each caUpdate In ActiveDocument.CoAuthoring.Updates
    strText = caUpdate.Range.Text
    MsgBox strText
Next caUpdate
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CoAuthUpdate object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained in the specified object. Read-only.
