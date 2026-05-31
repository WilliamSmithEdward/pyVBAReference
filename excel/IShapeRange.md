# IShapeRange

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002443B-0001-0000-C000-000000000046}  

## Properties (50)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Count As HRESULT  (read-only)`
- `_NewEnum As HRESULT  (read-only)`
- `Adjustments As HRESULT  (read-only)`
- `TextFrame As HRESULT  (read-only)`
- `AutoShapeType As HRESULT  (read/write)`
- `Callout As HRESULT  (read-only)`
- `ConnectionSiteCount As HRESULT  (read-only)`
- `Connector As HRESULT  (read-only)`
- `ConnectorFormat As HRESULT  (read-only)`
- `Fill As HRESULT  (read-only)`
- `GroupItems As HRESULT  (read-only)`
- `Height As HRESULT  (read/write)`
- `HorizontalFlip As HRESULT  (read-only)`
- `Left As HRESULT  (read/write)`
- `Line As HRESULT  (read-only)`
- `LockAspectRatio As HRESULT  (read/write)`
- `Name As HRESULT  (read/write)`
- `Nodes As HRESULT  (read-only)`
- `Rotation As HRESULT  (read/write)`
- `PictureFormat As HRESULT  (read-only)`
- `Shadow As HRESULT  (read-only)`
- `TextEffect As HRESULT  (read-only)`
- `ThreeD As HRESULT  (read-only)`
- `Top As HRESULT  (read/write)`
- `Type As HRESULT  (read-only)`
- `VerticalFlip As HRESULT  (read-only)`
- `Vertices As HRESULT  (read-only)`
- `Visible As HRESULT  (read/write)`
- `Width As HRESULT  (read/write)`
- `ZOrderPosition As HRESULT  (read-only)`
- `BlackWhiteMode As HRESULT  (read/write)`
- `AlternativeText As HRESULT  (read/write)`
- `Child As HRESULT  (read-only)`
- `ParentGroup As HRESULT  (read-only)`
- `ID As HRESULT  (read-only)`
- `Chart As HRESULT  (read-only)`
- `HasChart As HRESULT  (read-only)`
- `TextFrame2 As HRESULT  (read-only)`
- `ShapeStyle As HRESULT  (read/write)`
- `BackgroundStyle As HRESULT  (read/write)`
- `SoftEdge As HRESULT  (read-only)`
- `Glow As HRESULT  (read-only)`
- `Reflection As HRESULT  (read-only)`
- `Title As HRESULT  (read/write)`
- `GraphicStyle As HRESULT  (read/write)`
- `Model3D As HRESULT  (read-only)`
- `Decorative As HRESULT  (read/write)`

## Methods (21)

- `Item(Index As Variant, RHS As Shape)`
- `_Default(Index As Variant, RHS As Shape)`
- `Align(AlignCmd As MsoAlignCmd, RelativeTo As MsoTriState)`
- `Apply()`
- `Delete()`
- `Distribute(DistributeCmd As MsoDistributeCmd, RelativeTo As MsoTriState)`
- `Duplicate(RHS As ShapeRange)`
- `Flip(FlipCmd As MsoFlipCmd)`
- `IncrementLeft(Increment As Single)`
- `IncrementRotation(Increment As Single)`
- `IncrementTop(Increment As Single)`
- `Group(RHS As Shape)`
- `PickUp()`
- `RerouteConnections()`
- `Regroup(RHS As Shape)`
- `ScaleHeight(Factor As Single, RelativeToOriginalSize As MsoTriState, [Scale As Variant])`
- `ScaleWidth(Factor As Single, RelativeToOriginalSize As MsoTriState, [Scale As Variant])`
- `Select([Replace As Variant])`
- `SetShapesDefaultProperties()`
- `Ungroup(RHS As ShapeRange)`
- `ZOrder(ZOrderCmd As MsoZOrderCmd)`
