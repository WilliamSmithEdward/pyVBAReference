# Conflicts

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {C2B83A65-B061-4469-83B6-8877437CB8A0}  

A collection ofConflict objects that represents the conflicts in a document. The type of a Conflict object is specified by the WdRevisionType enumeration.

**Remarks:** Use the Conflicts property to return the Conflicts collection for a document. Use Conflicts (index), where index is the conflict index number, to return a single Conflict object.

**Example:**

```vba
ActiveDocument.CoAuthoring.Conflicts(1).Accept
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Conflicts object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of items in the Conflicts collection. Read-only.

## Methods (3)

- `Item(Index As Long) As Conflict`  
  Returns an individual Conflicts object in a collection.
    - `Index As Long` (required): The individual object to be returned.
- `AcceptAll()`  
  Accepts all of the user's changes, removes the conflicts, and merges the changes into the server copy of the document.
- `RejectAll()`  
  Rejects all of the user's changes and retains the server copy of the document.
