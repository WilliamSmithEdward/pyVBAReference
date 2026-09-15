# Frames

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002092B-0000-0000-C000-000000000046}  

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Frames object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of frames in the collection. Read-only.

## Methods (3)

- `Item(Index As Long) As Frame`  
  Returns an individual Frames object in a collection.
    - `Index As Long` (required): The individual object to be returned.
- `Add(Range As Range) As Frame`  
  Returns a Frame object that represents a new frame added to a range, selection, or document.
    - `Range As Range` (required): The range that you want the frame to surround.
- `Delete()`  
  Deletes the specified Frames collection.
