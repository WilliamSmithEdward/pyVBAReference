# CoAuthoring

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {65DF9F31-B1E3-4651-87E8-51D55F302161}  

Provides the primary entry point to the co authoring object model.

**Remarks:** The CoAuthoring object provides information about co authoring at the document level. For example, the CoAuthoring object can provide information about whether there are any locks in the document, which users have current locks in the document, or whether or not updates to the document content is available from the server. Use the CoAuthoring property to return the CoAuthoring object.

**Example:**

```vba
Sub CountLocks()
Dim i As Integer

i = ActiveDocument.CoAuthoring.Locks.Count

MsgBox i

End Sub
```

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CoAuthoring object.
- `Authors As CoAuthors  (read-only)`  
  Returns a CoAuthors collection that represents all the co authors currently editing the document. Read-only.
- `Me As CoAuthor  (read-only)`  
  Returns a CoAuthor object that represents the current user. Read-only.
- `PendingUpdates As Boolean  (read-only)`  
  Returns True if the document has pending updates that have not been accepted. Read-only.
- `Locks As CoAuthLocks  (read-only)`  
  Returns a CoAuthLocks collection that represents the locks in the document. Read-only.
- `Updates As CoAuthUpdates  (read-only)`  
  Returns a CoAuthUpdates collection that represents the most recent updates that were merged into the document. Read-only.
- `Conflicts As Conflicts  (read-only)`  
  Returns a Conflicts collection that represents all the conflicts in a document. Read-only.
- `CanShare As Boolean  (read-only)`  
  Returns a Boolean that specifies whether this document can be co authored. Read-only.
- `CanMerge As Boolean  (read-only)`  
  Returns a Boolean that specifies whether the document can be auto-merged. Read-only.
