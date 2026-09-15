# ProtectedViewWindows

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {BA72E559-4FF5-48F4-8215-5505F990966F}  

Represents a collection of ProtectedViewWindow objects.

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.

## Methods (2)

- `Item(Index As Long) As ProtectedViewWindow`  
  Returns a single ProtectedViewWindow object from the specified ProtectedViewWindows collection.
    - `Index As Long` (required): The index number of the single ProtectedViewWindow object in the collection to be returned.
- `Open(FileName As String, [ReadPassword As String], [OpenAndRepair As MsoTriState]) As ProtectedViewWindow`  
  Open and return a ProtectedViewWindow object from the ProtectedViewWindows collection.
    - `FileName As String` (required): The name of the file to open.
    - `ReadPassword As String` (optional): The password to use for the protected file.
    - `OpenAndRepair As MsoTriState` (optional): Indicates whether to repair the file.
