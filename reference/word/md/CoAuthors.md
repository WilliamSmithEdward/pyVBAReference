# CoAuthors

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {256B6ABA-6A38-4D39-971C-91FDA9922814}  

A collection of all the CoAuthor objects in the document.

**Remarks:** The CoAuthors collection contains all the co authors in the document (authors that are actively editing the document).

**Example:**

```vba
Dim i As Integer

i = ActiveDocument.CoAuthoring.Authors.Count

MsgBox "The number of co authors is " & i
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CoAuthors object.
- `Count As Long  (read-only)`  
  Returns the number of items in the CoAuthors collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As Variant) As CoAuthor`  
  Returns an individual CoAuthor object in a collection.
    - `Index As Variant` (required): The individual object to be returned.
