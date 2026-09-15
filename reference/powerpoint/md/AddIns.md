# AddIns

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493460-5A91-11CF-8700-00AA0060263B}  

A collection of AddIn objects that represent all the Microsoft PowerPoint-specific add-ins available to PowerPoint, regardless of whether or not they are loaded. This does not include Component Object Model (COM) add-ins.

**Example:**

```vba
For Each ad In AddIns

    If ad.Loaded Then MsgBox ad.Name

Next
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (3)

- `Item(Index As Variant) As AddIn`  
  Returns a single Addin object from the specified Addins collection.
    - `Index As Variant` (required): The name or index number of the single Addin object in the collection to be returned.
- `Add(FileName As String) As AddIn`  
  Returns an AddIn object that represents an add-in file added to the list of add-ins.
    - `FileName As String` (required): The full name of the file (including the path and file name extension) that contains the add-in you want to add to the list of add-ins.
- `Remove(Index As Variant)`  
  Removes an add-in from the collection of add-ins.
    - `Index As Variant` (required): The name or index number of the add-in to be removed from the collection.
