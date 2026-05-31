# SmartArtNode

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C8-0000-0000-C000-000000000046}  

A single semantic node within the data model of a SmartArt graphic.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(1).SmartArtNodes.Count
```

## Properties (11)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SmartArtNode object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SmartArtNode object was created. Read-only.
- `Parent As Object  (read-only)`  
  Returns the calling object. Read-only.
- `OrgChartLayout As MsoOrgChartLayoutType  (read/write)`  
  Retrieves or sets the MsoOrgChartLayoutType associated with this node if there is one. Read/write.
- `Shapes As ShapeRange  (read-only)`  
  Returns the shape range associated with this SmartArtNode object. Read-only.
- `TextFrame2 As TextFrame2  (read-only)`  
  Returns the text associated with the SmartArtNode object. Read-only.
- `Level As Long  (read-only)`  
  Retrieves the node's level in the hierarchy. Read-only.
- `Hidden As MsoTriState  (read-only)`  
  Returns True if this node is a hidden node in the data model. Read-only.
- `Nodes As SmartArtNodes  (read-only)`  
  Retrieves the children nodes associated with this SmartArt node. Read-only.
- `ParentNode As SmartArtNode  (read-only)`  
  Retrieves the parent SmartArtNode of this SmartArtNode. Read-only.
- `Type As MsoSmartArtNodeType  (read-only)`  
  Retrieves the type of SmartArt node. Read-only.

## Methods (8)

- `AddNode([Position As MsoSmartArtNodePosition], [Type As MsoSmartArtNodeType]) As SmartArtNode`  
  Adds a new SmartArtNode object to the data model in the way specified by the SmartArtNodePosition value, and of type SmartArtNodeType.
    - `Position As MsoSmartArtNodePosition` (optional): Specifies the location of the SmartArtNode in the data model; for example, msoSmartArtNodeAbove or msoSmartArtNodeAfter.
    - `Type As MsoSmartArtNodeType` (optional): Specifies the type of the added SmartArtNode; for example, msoSmartArtNodeTypeAssistant or msoSmartArtNodeTypeDefault.
- `Delete()`  
  Removes the current SmartArt node.
- `Promote()`  
  Promotes the current node (and all its children) a single level within the data model.
- `Demote()`  
  Demotes the current node a single level within the data model.
- `Larger()`  
  Increases the size of the SmartArt node. Mimics the behavior of the Larger button on the Microsoft Office Fluent Ribbon Format tab for SmartArt.
- `Smaller()`  
  Decreases the size of the SmartArt. Mimics the behavior of the Smaller button on the Microsoft Office Fluent Ribbon UI Format tab for SmartArt.
- `ReorderUp()`  
  Swaps a node with the previous node in the bulleted list. This method reorders the node's entire family.
- `ReorderDown()`  
  Swaps a node with the next node in the bulleted list. This method reorders the node's entire family.
