# DiagramNode

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934D8-5A91-11CF-8700-00AA0060263B}  

## Properties (9)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `Children As DiagramNodeChildren  (read-only)`
- `Shape As Shape  (read-only)`
- `Root As DiagramNode  (read-only)`
- `Diagram As Diagram  (read-only)`
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
