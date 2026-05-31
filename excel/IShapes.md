# IShapes

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002443A-0001-0000-C000-000000000046}  

## Properties (6)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Count As HRESULT  (read-only)`
- `_NewEnum As HRESULT  (read-only)`
- `Range As HRESULT  (read-only)`

## Methods (20)

- `Item(Index As Variant, RHS As Shape)`
- `_Default(Index As Variant, RHS As Shape)`
- `AddCallout(Type As MsoCalloutType, Left As Single, Top As Single, Width As Single, Height As Single, RHS As Shape)`
- `AddConnector(Type As MsoConnectorType, BeginX As Single, BeginY As Single, EndX As Single, EndY As Single, RHS As Shape)`
- `AddCurve(SafeArrayOfPoints As Variant, RHS As Shape)`
- `AddLabel(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single, RHS As Shape)`
- `AddLine(BeginX As Single, BeginY As Single, EndX As Single, EndY As Single, RHS As Shape)`
- `AddPicture(Filename As String, LinkToFile As MsoTriState, SaveWithDocument As MsoTriState, Left As Single, Top As Single, Width As Single, Height As Single, RHS As Shape)`
- `AddPolyline(SafeArrayOfPoints As Variant, RHS As Shape)`
- `AddShape(Type As MsoAutoShapeType, Left As Single, Top As Single, Width As Single, Height As Single, RHS As Shape)`
- `AddTextEffect(PresetTextEffect As MsoPresetTextEffect, Text As String, FontName As String, FontSize As Single, FontBold As MsoTriState, FontItalic As MsoTriState, Left As Single, Top As Single, RHS As Shape)`
- `AddTextbox(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single, RHS As Shape)`
- `BuildFreeform(EditingType As MsoEditingType, X1 As Single, Y1 As Single, RHS As FreeformBuilder)`
- `SelectAll()`
- `AddFormControl(Type As XlFormControl, Left As Long, Top As Long, Width As Long, Height As Long, RHS As Shape)`
- `AddOLEObject([ClassType As Variant], [Filename As Variant], [Link As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], RHS As Shape)`
- `AddSmartArt(Layout As SmartArtLayout, [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], RHS As Shape)`
- `AddChart2([Style As Variant], [XlChartType As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], [NewLayout As Variant], RHS As Shape)`
- `AddPicture2(Filename As String, LinkToFile As MsoTriState, SaveWithDocument As MsoTriState, Left As Single, Top As Single, Width As Single, Height As Single, Compress As MsoPictureCompress, RHS As Shape)`
- `Add3DModel(Filename As String, [LinkToFile As Variant], [SaveWithDocument As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], RHS As Shape)`
