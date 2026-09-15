# Words

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002095C-0000-0000-C000-000000000046}  

A collection of words in a selection, range, or document. Each item in the Words collection is a Range object that represents one word. There is no Word object.

**Remarks:** Use the Words property to return the Words object. The following code example displays how many words are currently selected. Use Words (Index), where Index is the index number, to return a Range object that represents one word. The index number represents the position of the word in the Words collection. The following code example formats the first word in the selection as 24-point italic. The item in the Words collection includes both the word and the spaces after the word. To remove the trailing spaces, use the Visual Basic RTrim function - for example, _RTrim(ActiveDocument.Words(1))_. The following code example selects the first word (and its trailing spaces) in the active document. If the selection is the insertion point and it is immediately followed by a space, _Selection.Words(1)_ refers to the word preceding the selection. If the selection is the insertion point and is immediately followed by a character, _Selection.Words(1)_ refers to the word following the selection. The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object.

## Properties (7)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of words in the collection. Read-only.
- `First As Range  (read-only)`  
  Returns a Range object that represents the first word in a collection of words.
- `Last As Range  (read-only)`  
  Returns a Range object that represents the last word in a collection of words.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Words object.

## Methods (1)

- `Item(Index As Long) As Range`  
  Returns an individual Range object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
