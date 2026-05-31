# SmartArtNodes

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C7-0000-0000-C000-000000000046}  

Represents a collection of nodes within a SmartArt diagram.

**Remarks:** These nodes correspond directly to semantic elements contained within the data model of the graphic.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(1).SmartArtNodes.Count
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SmartArtNodes object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SmartArtNodes object was created. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Parent As Object  (read-only)`  
  Returns the calling object. Read-only.
- `Count As Long  (read-only)`  
  Retrieves the number of SmartArtNode objects contained within the SmartArtNodes collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As SmartArtNode`  
  Retrieves a SmartArtNode object at the specified index or with the specified unique Id.
    - `Index As Variant` (required): Specifies either an integer representing the index or a string representing the location of the SmartArtNode object.
- `Add() As SmartArtNode`  
  Adds a new SmartArtNode object to the diagram with specified text.
