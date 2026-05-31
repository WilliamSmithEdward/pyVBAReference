# Drawings

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208A9-0000-0000-C000-000000000046}  

## Properties (31)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Enabled As Boolean  (read/write)`
- `Height As Double  (read/write)`
- `Left As Double  (read/write)`
- `Locked As Boolean  (read/write)`
- `Placement As Variant  (read/write)`
- `PrintObject As Boolean  (read/write)`
- `Top As Double  (read/write)`
- `Visible As Boolean  (read/write)`
- `Width As Double  (read/write)`
- `ZOrder As Long  (read-only)`
- `ShapeRange As ShapeRange  (read-only)`
- `AddIndent As Boolean  (read/write)`
- `AutoScaleFont As Variant  (read/write)`
- `AutoSize As Boolean  (read/write)`
- `Caption As String  (read/write)`
- `Characters As Characters  (read-only)`
- `Font As Font  (read-only)`
- `Formula As String  (read/write)`
- `HorizontalAlignment As Variant  (read/write)`
- `LockedText As Boolean  (read/write)`
- `Orientation As Variant  (read/write)`
- `Text As String  (read/write)`
- `VerticalAlignment As Variant  (read/write)`
- `ReadingOrder As Long  (read/write)`
- `Border As Border  (read-only)`
- `Interior As Interior  (read-only)`
- `Shadow As Boolean  (read/write)`
- `Count As Long  (read-only)`

## Methods (14)

- `BringToFront() As Variant`
- `Copy() As Variant`
- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat]) As Variant`
- `Cut() As Variant`
- `Delete() As Variant`
- `Duplicate() As Object`
- `Select([Replace As Variant]) As Variant`
- `SendToBack() As Variant`
- `CheckSpelling([CustomDictionary As Variant], [IgnoreUppercase As Variant], [AlwaysSuggest As Variant], [SpellLang As Variant]) As Variant`
- `Reshape(Vertex As Long, Insert As Boolean, [Left As Variant], [Top As Variant]) As Variant`
- `Add(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Closed As Boolean) As Drawing`
- `Group() As GroupObject`
- `Item(Index As Variant) As Object`
- `_NewEnum() As IUnknown`
