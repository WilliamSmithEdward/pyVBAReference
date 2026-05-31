# SmartArtColors

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03CD-0000-0000-C000-000000000046}  

A collection of SmartArtColor objects.

**Remarks:** Simulates the commands on the Microsoft Office Fluent Ribbon user interface on the SmartArt Tools tab, on the Design group, and on the Change Colors command.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SmartArtColors object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SmartArtColors object was created. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Parent As Object  (read-only)`  
  Returns the calling object. Read-only.
- `Count As Long  (read-only)`  
  Retrieves the count of the number of SmartArtColor objects contained within the SmartArtColors collection. Read-only.

## Methods (1)

- `Item(Index As Variant) As SmartArtColor`  
  Retrieves a SmartArtColor object at the specified index or with the specified unique Id.
    - `Index As Variant` (required): Specifies either an integer representing the index or a string representing the Id of the SmartArt color.
