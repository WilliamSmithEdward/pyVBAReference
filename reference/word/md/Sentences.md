# Sentences

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002095B-0000-0000-C000-000000000046}  

A collection of Range objects that represent all the sentences in a selection, range, or document. There is no Sentence object.

**Remarks:** Use the Sentences property to return the Sentences collection. The following example displays the number of sentences selected. Use Sentences (Index), where Index is the index number, to return a Range object that represents a sentence. The index number represents the position of a sentence in the Sentences collection. The following example formats the first sentence in the active document. The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object. The Add method isn't available for the Sentences collection. Instead, use the InsertAfter or InsertBefore method to add a sentence to a Range object. The following example inserts a sentence after the first sentence in the active document.

## Properties (7)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of sentences in the collection. Read-only.
- `First As Range  (read-only)`  
  Returns a Range object that represents the first sentence in a collection of sentences within a document, range, or selection.
- `Last As Range  (read-only)`  
  Returns a Range object that represents the last sentence in a document, selection, or range.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Sentences object.

## Methods (1)

- `Item(Index As Long) As Range`  
  Returns an individual Range object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
