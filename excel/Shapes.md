# Shapes

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002443A-0000-0000-C000-000000000046}  

## Properties (6)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Range As ShapeRange  (read-only)`

## Methods (20)

- `Item(Index As Variant) As Shape`
- `_Default(Index As Variant) As Shape`
- `AddCallout(Type As MsoCalloutType, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`
- `AddConnector(Type As MsoConnectorType, BeginX As Single, BeginY As Single, EndX As Single, EndY As Single) As Shape`
- `AddCurve(SafeArrayOfPoints As Variant) As Shape`
- `AddLabel(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`
- `AddLine(BeginX As Single, BeginY As Single, EndX As Single, EndY As Single) As Shape`
- `AddPicture(Filename As String, LinkToFile As MsoTriState, SaveWithDocument As MsoTriState, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`
- `AddPolyline(SafeArrayOfPoints As Variant) As Shape`
- `AddShape(Type As MsoAutoShapeType, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`
- `AddTextEffect(PresetTextEffect As MsoPresetTextEffect, Text As String, FontName As String, FontSize As Single, FontBold As MsoTriState, FontItalic As MsoTriState, Left As Single, Top As Single) As Shape`
- `AddTextbox(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`
- `BuildFreeform(EditingType As MsoEditingType, X1 As Single, Y1 As Single) As FreeformBuilder`
- `SelectAll()`
- `AddFormControl(Type As XlFormControl, Left As Long, Top As Long, Width As Long, Height As Long) As Shape`
- `AddOLEObject([ClassType As Variant], [Filename As Variant], [Link As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As Shape`
- `AddSmartArt(Layout As SmartArtLayout, [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As Shape`
- `AddChart2([Style As Variant], [XlChartType As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], [NewLayout As Variant]) As Shape`
- `AddPicture2(Filename As String, LinkToFile As MsoTriState, SaveWithDocument As MsoTriState, Left As Single, Top As Single, Width As Single, Height As Single, Compress As MsoPictureCompress) As Shape`
- `Add3DModel(Filename As String, [LinkToFile As Variant], [SaveWithDocument As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As Shape`
