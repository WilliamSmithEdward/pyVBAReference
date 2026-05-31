# SmartArt

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C6-0000-0000-C000-000000000046}  

The top level class for interacting with a SmartArt graphic.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(1).SmartArt.Nodes.Add
```

## Properties (9)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SmartArt object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SmartArt object was created. Read-only.
- `Parent As Object  (read-only)`  
  Returns the calling object. Read-only.
- `AllNodes As SmartArtNodes  (read-only)`  
  Retrieves a SmartArtNodes object containing all of the nodes within the SmartArt diagram. Read-only.
- `Nodes As SmartArtNodes  (read-only)`  
  Retrieves the children of the root node of the SmartArt diagram. Read-only.
- `Layout As SmartArtLayout  (read/write)`  
  Retrieves or sets the SmartArt layout associated with the SmartArt graphic. Read/write.
- `QuickStyle As SmartArtQuickStyle  (read/write)`  
  Retrieves or sets the SmartArt quick style applied to the SmartArt graphic. Read/write.
- `Color As SmartArtColor  (read/write)`  
  Retrieves or sets the SmartArt color style applied to the SmartArt graphic. Read/write.
- `Reverse As MsoTriState  (read/write)`  
  Gets or sets the state of the SmartArt diagram with LTR (left-to-right) or RTL (right-to-left), if the diagram supports reversal. Read/write.

## Methods (1)

- `Reset()`  
  Resets the SmartArt graphic to its original state.
