# Shape

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C031C-0000-0000-C000-000000000046}  

## Properties (51)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `Adjustments As Adjustments  (read-only)`
- `AutoShapeType As MsoAutoShapeType  (read/write)`
- `BlackWhiteMode As MsoBlackWhiteMode  (read/write)`
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
- `TextFrame As TextFrame  (read-only)`
- `ThreeD As ThreeDFormat  (read-only)`
- `Top As Single  (read/write)`
- `Type As MsoShapeType  (read-only)`
- `VerticalFlip As MsoTriState  (read-only)`
- `Vertices As Variant  (read-only)`
- `Visible As MsoTriState  (read/write)`
- `Width As Single  (read/write)`
- `ZOrderPosition As Long  (read-only)`
- `AlternativeText As String  (read/write)`
- `Child As MsoTriState  (read-only)`
- `ParentGroup As Shape  (read-only)`
- `Id As Long  (read-only)`
- `TextFrame2 As TextFrame2  (read-only)`
- `HasChart As MsoTriState  (read-only)`
- `Chart As IMsoChart  (read-only)`
- `ShapeStyle As MsoShapeStyleIndex  (read/write)`
- `BackgroundStyle As MsoBackgroundStyleIndex  (read/write)`
- `SoftEdge As SoftEdgeFormat  (read-only)`
- `Glow As GlowFormat  (read-only)`
- `Reflection As ReflectionFormat  (read-only)`
- `HasSmartArt As MsoTriState  (read-only)`
- `SmartArt As SmartArt  (read-only)`
- `Title As String  (read/write)`
- `GraphicStyle As MsoGraphicStyleIndex  (read/write)`
- `Model3D As Model3DFormat  (read-only)`
- `Decorative As MsoTriState  (read/write)`
- `Locked As MsoTriState  (read/write)`

## Methods (19)

- `Apply()`
- `Delete()`
- `Duplicate() As Shape`
- `Flip(FlipCmd As MsoFlipCmd)`
- `IncrementLeft(Increment As Single)`
- `IncrementRotation(Increment As Single)`
- `IncrementTop(Increment As Single)`
- `PickUp()`
- `RerouteConnections()`
- `ScaleHeight(Factor As Single, RelativeToOriginalSize As MsoTriState, [fScale As MsoScaleFrom])`
- `ScaleWidth(Factor As Single, RelativeToOriginalSize As MsoTriState, [fScale As MsoScaleFrom])`
- `Select([Replace As Variant])`
- `SetShapesDefaultProperties()`
- `Ungroup() As ShapeRange`
- `ZOrder(ZOrderCmd As MsoZOrderCmd)`
- `Cut()`
- `Copy()`
- `ConvertTextToSmartArt(Layout As SmartArtLayout)`
- `SaveAsPicture(PictureType As MsoPictureType, FileName As String, FSaveShapesIndividually As Boolean)`
