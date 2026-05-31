# SmartArtQuickStyles

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03CB-0000-0000-C000-000000000046}  

Represents a collection of SmartArtQuickStyle objects.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles(i)
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SmartArtQuickStyles object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SmartArtQuickStyles object was created. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Parent As Object  (read-only)`  
  Returns the calling object. Read-only.
- `Count As Long  (read-only)`  
  Retrieves the count of the number of SmartArtQuickStyle objects contained within the SmartArtQuickStyles collection. Read-only.

## Methods (1)

- `Item(Index As Variant) As SmartArtQuickStyle`  
  Retrieves a SmartArtQuickStyle object at the specified index or with the specified unique Id.
    - `Index As Variant` (required): Specifies either an integer representing the index or a string representing the location of the SmartArtQuickStyle object.
