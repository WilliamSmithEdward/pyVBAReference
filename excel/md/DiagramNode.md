# DiagramNode

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C0370-0000-0000-C000-000000000046}  

## Properties (9)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `Children As DiagramNodeChildren  (read-only)`
- `Shape As Shape  (read-only)`
- `Root As DiagramNode  (read-only)`
- `Diagram As IMsoDiagram  (read-only)`
- `Layout As MsoOrgChartLayoutType  (read/write)`
- `TextShape As Shape  (read-only)`

## Methods (9)

- `AddNode([pos As MsoRelativeNodePosition], [nodeType As MsoDiagramNodeType]) As DiagramNode`
- `Delete()`
- `MoveNode(pTargetNode As DiagramNode, pos As MsoRelativeNodePosition)`
- `ReplaceNode(pTargetNode As DiagramNode)`
- `SwapNode(pTargetNode As DiagramNode, [swapChildren As Boolean])`
- `CloneNode(copyChildren As Boolean, pTargetNode As DiagramNode, [pos As MsoRelativeNodePosition]) As DiagramNode`
- `TransferChildren(pReceivingNode As DiagramNode)`
- `NextNode() As DiagramNode`
- `PrevNode() As DiagramNode`
