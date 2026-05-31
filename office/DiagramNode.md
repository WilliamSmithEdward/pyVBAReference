# DiagramNode

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
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

- `AddNode([Pos As MsoRelativeNodePosition], [NodeType As MsoDiagramNodeType]) As DiagramNode`
- `Delete()`
- `MoveNode(TargetNode As DiagramNode, Pos As MsoRelativeNodePosition)`
- `ReplaceNode(TargetNode As DiagramNode)`
- `SwapNode(TargetNode As DiagramNode, [SwapChildren As Boolean])`
- `CloneNode(CopyChildren As Boolean, TargetNode As DiagramNode, [Pos As MsoRelativeNodePosition]) As DiagramNode`
- `TransferChildren(ReceivingNode As DiagramNode)`
- `NextNode() As DiagramNode`
- `PrevNode() As DiagramNode`
