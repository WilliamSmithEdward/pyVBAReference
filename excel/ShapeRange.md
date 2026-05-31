# ShapeRange

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002443B-0000-0000-C000-000000000046}  

## Properties (50)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Adjustments As Adjustments  (read-only)`
- `TextFrame As TextFrame  (read-only)`
- `AutoShapeType As MsoAutoShapeType  (read/write)`
- `Callout As CalloutFormat  (read-only)`
- `ConnectionSiteCount As Long  (read-only)`
- `Connector As MsoTriState  (read-only)`
- `ConnectorFormat As ConnectorFormat  (read-only)`
- `Fill As FillFormat  (read-only)`
- `GroupItems As GroupShapes  (read-only)`
- `Height As Single  (read/write)`
- `HorizontalFlip As MsoTriState  (read-only)`
- `Left As Single  (read/write)`
- `Line As LineFormat  (read-only)`
- `LockAspectRatio As MsoTriState  (read/write)`
- `Name As String  (read/write)`
- `Nodes As ShapeNodes  (read-only)`
- `Rotation As Single  (read/write)`
- `PictureFormat As PictureFormat  (read-only)`
- `Shadow As ShadowFormat  (read-only)`
- `TextEffect As TextEffectFormat  (read-only)`
- `ThreeD As ThreeDFormat  (read-only)`
- `Top As Single  (read/write)`
- `Type As MsoShapeType  (read-only)`
- `VerticalFlip As MsoTriState  (read-only)`
- `Vertices As Variant  (read-only)`
- `Visible As MsoTriState  (read/write)`
- `Width As Single  (read/write)`
- `ZOrderPosition As Long  (read-only)`
- `BlackWhiteMode As MsoBlackWhiteMode  (read/write)`
- `AlternativeText As String  (read/write)`
- `Child As MsoTriState  (read-only)`
- `ParentGroup As Shape  (read-only)`
- `ID As Long  (read-only)`
- `Chart As Chart  (read-only)`
- `HasChart As MsoTriState  (read-only)`
- `TextFrame2 As TextFrame2  (read-only)`
- `ShapeStyle As MsoShapeStyleIndex  (read/write)`
- `BackgroundStyle As MsoBackgroundStyleIndex  (read/write)`
- `SoftEdge As SoftEdgeFormat  (read-only)`
- `Glow As GlowFormat  (read-only)`
- `Reflection As ReflectionFormat  (read-only)`
- `Title As String  (read/write)`
- `GraphicStyle As MsoGraphicStyleIndex  (read/write)`
- `Model3D As Model3DFormat  (read-only)`
- `Decorative As MsoTriState  (read/write)`

## Methods (21)

- `Item(Index As Variant) As Shape`
- `_Default(Index As Variant) As Shape`
- `Align(AlignCmd As MsoAlignCmd, RelativeTo As MsoTriState)`
- `Apply()`
- `Delete()`
- `Distribute(DistributeCmd As MsoDistributeCmd, RelativeTo As MsoTriState)`
- `Duplicate() As ShapeRange`
- `Flip(FlipCmd As MsoFlipCmd)`
- `IncrementLeft(Increment As Single)`
- `IncrementRotation(Increment As Single)`
- `IncrementTop(Increment As Single)`
- `Group() As Shape`
- `PickUp()`
- `RerouteConnections()`
- `Regroup() As Shape`
- `ScaleHeight(Factor As Single, RelativeToOriginalSize As MsoTriState, [Scale As Variant])`
- `ScaleWidth(Factor As Single, RelativeToOriginalSize As MsoTriState, [Scale As Variant])`
- `Select([Replace As Variant])`
- `SetShapesDefaultProperties()`
- `Ungroup() As ShapeRange`
- `ZOrder(ZOrderCmd As MsoZOrderCmd)`
