# SmartArtLayouts

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C9-0000-0000-C000-000000000046}  

Represents a collection of SmartArt layout diagrams.

**Remarks:** Choices include Basic Block list, Picture Caption list, Vertical Bulleted list, etc.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SmartArtLayouts object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SmartArtLayouts object was created. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Parent As Object  (read-only)`  
  Returns the calling object. Read-only.
- `Count As Long  (read-only)`  
  Retrieves the count of the number of SmartArtLayout objects contained within the SmartArtLayouts collection. Read-only.

## Methods (1)

- `Item(Index As Variant) As SmartArtLayout`  
  Retrieves a SmartArtLayout object at the specified index or with the specified unique Id.
    - `Index As Variant` (required): Specifies either an integer representing the index or a string representing the location of the SmartArtLayout object.
