# Revisions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020980-0000-0000-C000-000000000046}  

A collection of Revision objects that represent the changes marked with revision marks in a range or document.

**Remarks:** Use the Revisions property to return the Revisions collection. The following code example displays the number of revisions in the main text story. The following code example accepts all the revisions in the selection. The following code example accepts all the revisions in the first paragraph in the selection. The Add method is not available for the Revisions collection. Revision objects are added when change tracking is enabled. Set the TrackRevisions property to True to track revisions made to the document text. The following code example enables revision tracking in the active document and then inserts "The " before the selection. Use Revisions (Index), where Index is the index number, to return a single Revision object. The index number represents the position of the revision in the range or document. The following code example displays the author name for the first revision in the first section. The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Revisions object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of revisions in the collection. Read-only.

## Methods (3)

- `Item(Index As Long) As Revision`  
  Returns an individual Revisions object in a collection.
    - `Index As Long` (required): The individual object to be returned.
- `AcceptAll()`  
  Accepts all the tracked changes in a document or range, removes all revision marks, and incorporates the changes into the document.
- `RejectAll()`
