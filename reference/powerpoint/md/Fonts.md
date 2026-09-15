# Fonts

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493467-5A91-11CF-8700-00AA0060263B}  

A collection of all the Font objects in the specified presentation.

**Remarks:** Each Font object represents a font that's used in the presentation.

**Example:**

```vba
MsgBox ActivePresentation.Fonts.Count
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (2)

- `Item(Index As Variant) As Font`  
  Returns a single Font object from the specified Fonts collection.
    - `Index As Variant` (required): The name or index number of the single Font object in the collection to be returned.
- `Replace(Original As String, Replacement As String)`  
  Replaces a font in the Fonts collection.
    - `Original As String` (required): The name of the font to replace.
    - `Replacement As String` (required): The name of the replacement font.
