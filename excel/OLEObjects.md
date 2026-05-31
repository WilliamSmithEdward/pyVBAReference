# OLEObjects

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208A3-0000-0000-C000-000000000046}  

## Properties (20)

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
- `Border As Border  (read-only)`
- `Interior As Interior  (read-only)`
- `Shadow As Boolean  (read/write)`
- `AutoLoad As Boolean  (read/write)`
- `SourceName As String  (read/write)`
- `Count As Long  (read-only)`

## Methods (11)

- `BringToFront() As Variant`
- `Copy() As Variant`
- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat]) As Variant`
- `Cut() As Variant`
- `Delete() As Variant`
- `Duplicate() As Object`
- `Select([Replace As Variant]) As Variant`
- `SendToBack() As Variant`
- `Add([ClassType As Variant], [Filename As Variant], [Link As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As OLEObject`
- `Item(Index As Variant) As Object`
- `_NewEnum() As IUnknown`
