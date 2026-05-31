# ChartObjects

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208D0-0000-0000-C000-000000000046}  

## Properties (14)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Height As Double  (read/write)`
- `Left As Double  (read/write)`
- `Locked As Boolean  (read/write)`
- `Placement As Variant  (read/write)`
- `PrintObject As Boolean  (read/write)`
- `Top As Double  (read/write)`
- `Visible As Boolean  (read/write)`
- `Width As Double  (read/write)`
- `ShapeRange As ShapeRange  (read-only)`
- `ProtectChartObject As Boolean  (read/write)`
- `Count As Long  (read-only)`

## Methods (10)

- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat]) As Variant`
- `Cut() As Variant`
- `Delete() As Variant`
- `Duplicate() As Object`
- `Select([Replace As Variant]) As Variant`
- `Copy() As Variant`
- `Add(Left As Double, Top As Double, Width As Double, Height As Double) As ChartObject`
- `Item(Index As Variant) As Object`
- `_NewEnum() As IUnknown`
- `_Default(Index As Variant) As Object`
