# Reviewers

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {12DCDC9A-5418-48A3-BBE6-EB57BAE275E8}  

A collection of Reviewer objects that represents the reviewers of one or more documents. The Reviewers collection contains the names of all reviewers who have reviewed documents opened or edited on a computer.

**Remarks:** Use Reviewers (Index), where Index is the name or index number of the reviewer, to return a single reviewer in the Reviewers collection. This example hides revisions made by the first reviewer in the Reviewers collection.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Reviewers object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of reviewers in the collection. Read-only.

## Methods (1)

- `Item(Index As Variant) As Reviewer`  
  Returns an individual Reviewer object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
