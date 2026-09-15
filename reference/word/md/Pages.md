# Pages

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {91807402-6C6F-47CD-B8FA-C42FEE8EE924}  

A collection of pages in a document. Use the Pages collection and the related objects and properties for programmatically defining page layout in a document.

**Remarks:** Use the Pages property to return a Pages collection. The following example accesses all pages in the active document. Use the Item method to access an individual Page object that represents an individual page in a document. The following example accesses the first page in the active document.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of pages in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Pages object.

## Methods (1)

- `Item(Index As Long) As Page`  
  Returns an individual Page object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
