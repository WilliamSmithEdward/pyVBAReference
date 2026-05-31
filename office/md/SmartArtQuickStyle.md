# SmartArtQuickStyle

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03CC-0000-0000-C000-000000000046}  

Represents a SmartArt quick style.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles(i)
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SmartArtQuickStyle object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SmartArtQuickStyle object was created. Read-only.
- `Parent As Object  (read-only)`  
  Returns the calling object. Read-only.
- `Id As String  (read-only)`  
  Retrieves the unique Id of the associated SmartArtQuickStyle object. Read-only.
- `Name As String  (read-only)`  
  Retrieves the string name of the SmartArtQuickStyle object. Read-only.
- `Description As String  (read-only)`  
  Retrieves the description of the SmartArtQuickStyle object. Read-only.
- `Category As String  (read-only)`  
  Retrieves the primary category name associated with the SmartArtQuickStyle object. Read-only.
