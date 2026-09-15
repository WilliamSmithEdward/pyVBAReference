# Characters

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002095D-0000-0000-C000-000000000046}  

A collection of characters in a selection, range, or document. There is no Character object; instead, each item in the Characters collection is a Range object that represents one character.

**Remarks:** Use the Characters property of a Document, Range, or Selection object to return the Characters collection. The following example displays how many characters are selected. Use Characters (Index), where Index is the index number, to return a Range object that represents one character. The index number represents the position of a character in the Characters collection. The following example formats the first letter in the selection as 24-point bold. Remarks The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object. There is no Add method for the Characters collection. Instead, use the InsertAfter or InsertBefore method to add characters to a Range object. The following example inserts a new paragraph after the first paragraph in the active document.

## Properties (7)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of characters in the Characters collection. Read-only Long.
- `First As Range  (read-only)`  
  Returns a Range object that represents the first character in the collection.
- `Last As Range  (read-only)`  
  Returns a Range object that represents the last character in the collection.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Characters object.

## Methods (1)

- `Item(Index As Long) As Range`  
  Returns an individual Range object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
