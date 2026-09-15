# CoAuthUpdates

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {30225CFC-5A71-4FE6-B527-90A52C54AE77}  

A collection of CoAuthUpdate objects that represent the updates that were merged into the document at the last explicit save.

**Remarks:** When a document with co authoring enabled is edited by more than one author, changes to the document by one author are pushed to other authors' versions of the document using updates. When a co author performs an explicit document save (by pressing CTRL + S, for example), changes made by other co authors are merged into the document as updates. The CoAuthUpdates collection contains all the changes that were merged into the document, where each change is a single update. The contents of the CoAuthUpdates collection remains the same until a co author performs another explicit document save. When the co author saves the document again, if there are no new changes from other co authors that are merged into the document, the CoAuthUpdates collection retains the same updates that were merged at the previous explicit save. If there are new changes that are merged into the document, the CoAuthUpdates collection contains the new updates for the document.

**Example:**

```vba
Dim countOfUpdates As Integer

countOfUpdates = ActiveDocument.CoAuthoring.Updates.Count

MsgBox "The number of updates is " & countOfUpdates
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CoAuthUpdates object.
- `Count As Long  (read-only)`  
  Returns the number of items in the CoAuthUpdates collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As Long) As CoAuthUpdate`  
  Returns an individual CoAuthUpdates object in a collection.
    - `Index As Long` (required): The individual object to be returned.
