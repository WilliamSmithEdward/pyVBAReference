# CanvasShapes

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0371-0000-0000-C000-000000000046}  

## Properties (6)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Background As Shape  (read-only)`

## Methods (14)

- `Item(Index As Variant) As Shape`
- `AddCallout(Type As MsoCalloutType, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`
- `AddConnector(Type As MsoConnectorType, BeginX As Single, BeginY As Single, EndX As Single, EndY As Single) As Shape`
- `AddCurve(SafeArrayOfPoints As Variant) As Shape`
- `AddLabel(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`
- `AddLine(BeginX As Single, BeginY As Single, EndX As Single, EndY As Single) As Shape`
- `AddPicture(FileName As String, LinkToFile As MsoTriState, SaveWithDocument As MsoTriState, Left As Single, Top As Single, [Width As Single], [Height As Single]) As Shape`
- `AddPolyline(SafeArrayOfPoints As Variant) As Shape`
- `AddShape(Type As MsoAutoShapeType, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`
- `AddTextEffect(PresetTextEffect As MsoPresetTextEffect, Text As String, FontName As String, FontSize As Single, FontBold As MsoTriState, FontItalic As MsoTriState, Left As Single, Top As Single) As Shape`
- `AddTextbox(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`
- `BuildFreeform(EditingType As MsoEditingType, X1 As Single, Y1 As Single) As FreeformBuilder`
- `Range(Index As Variant) As ShapeRange`
- `SelectAll()`
