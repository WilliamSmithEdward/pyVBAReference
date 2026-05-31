# SmartArtColor

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03CE-0000-0000-C000-000000000046}  

Chooses the color scheme for the SmartArt diagram.

**Remarks:** Simulates the commands on the Microsoft Office Fluent Ribbon user interface on the SmartArt Tools tab, on the Design group, and on the Change Colors command.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SmartArtColor object. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only.
- `Parent As Object  (read-only)`  
  Returns the calling object. Read-only.
- `Id As String  (read-only)`  
  Retrieves the unique Id of the associated SmartArt color style. Read-only.
- `Name As String  (read-only)`  
  Retrieves the string name of the SmartArt color style. Read-only.
- `Description As String  (read-only)`  
  Retrieves the description of the SmartArt color style. Read-only.
- `Category As String  (read-only)`  
  Retrieves the primary category name associated with the SmartArt color style. Read-only.
