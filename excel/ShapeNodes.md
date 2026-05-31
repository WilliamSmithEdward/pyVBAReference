# ShapeNodes

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C0319-0000-0000-C000-000000000046}  

## Properties (5)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (6)

- `Item(Index As Variant) As ShapeNode`
- `Delete(Index As Long)`
- `Insert(Index As Long, SegmentType As MsoSegmentType, EditingType As MsoEditingType, X1 As Single, Y1 As Single, [X2 As Single], [Y2 As Single], [X3 As Single], [Y3 As Single])`
- `SetEditingType(Index As Long, EditingType As MsoEditingType)`
- `SetPosition(Index As Long, X1 As Single, Y1 As Single)`
- `SetSegmentType(Index As Long, SegmentType As MsoSegmentType)`
