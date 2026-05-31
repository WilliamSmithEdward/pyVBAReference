# SmartArtLayout

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03CA-0000-0000-C000-000000000046}  

Represents a SmartArt diagram.

**Remarks:** Choices include Basic Block list, Picture Caption list, Vertical Bulleted list, etc.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SmartArtLayout object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SmartArtLayout object was created. Read-only.
- `Parent As Object  (read-only)`  
  Returns the calling object. Read-only.
- `Id As String  (read-only)`  
  Retrieves the unique Id of the associated SmartArt layout. Read-only.
- `Name As String  (read-only)`  
  Retrieves the string name of the SmartArt layout. Read-only.
- `Description As String  (read-only)`  
  Retrieves the description of the SmartArt layout. Read-only.
- `Category As String  (read-only)`  
  Retrieves the primary category name associated with the SmartArt layout. Read-only.
