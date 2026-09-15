# Sources

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {FA02A26B-6550-45C5-B6F0-80E757CD3482}  

Represents a collection of Source objects.

**Remarks:** For more information, see Working with Bibliographies.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of items in the Sources collection. Read-only Long.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Sources object.

## Methods (2)

- `Item(Index As Long) As Source`  
  Returns a Source object that represents the specified item in the collection.
    - `Index As Long` (required): Specifies the ordinal position of the object within the collection.
- `Add(Data As String)`  
  Add a new source to the sources listed in the Source Manager dialog box.
    - `Data As String` (required): An XML string that represents the field values for the new source.
