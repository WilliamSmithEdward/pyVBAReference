# SmartArtNode

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C8-0000-0000-C000-000000000046}  

## Properties (11)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `OrgChartLayout As MsoOrgChartLayoutType  (read/write)`
- `Shapes As ShapeRange  (read-only)`
- `TextFrame2 As TextFrame2  (read-only)`
- `Level As Long  (read-only)`
- `Hidden As MsoTriState  (read-only)`
- `Nodes As SmartArtNodes  (read-only)`
- `ParentNode As SmartArtNode  (read-only)`
- `Type As MsoSmartArtNodeType  (read-only)`

## Methods (8)

- `AddNode([Position As MsoSmartArtNodePosition], [Type As MsoSmartArtNodeType]) As SmartArtNode`
- `Delete()`
- `Promote()`
- `Demote()`
- `Larger()`
- `Smaller()`
- `ReorderUp()`
- `ReorderDown()`
