# Shapes

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002443A-0000-0000-C000-000000000046}  

A collection of all the Shape objects on the specified sheet.

**Remarks:** Each Shape object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.

**Example:**

```vba
Set myDocument = Worksheets(1)
myDocument.Shapes.SelectAll
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `_NewEnum As IUnknown  (read-only)`
- `Range As ShapeRange  (read-only)`  
  Returns a ShapeRange object that represents a subset of the shapes in a Shapes collection.

## Methods (20)

- `Item(Index As Variant) As Shape`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_Default(Index As Variant) As Shape`
- `AddCallout(Type As MsoCalloutType, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Creates a borderless line callout. Returns a Shape object that represents the new callout.
    - `Type As MsoCalloutType` (required): The type of callout line.
    - `Left As Single` (required): The position (in points) of the upper-left corner of the callout's bounding box relative to the upper-left corner of the document.
    - `Top As Single` (required): The position (in points) of the upper-left corner of the callout's bounding box relative to the top of the document.
    - `Width As Single` (required): The width of the callout's bounding box, in points.
    - `Height As Single` (required): The height of the callout's bounding box, in points.
- `AddConnector(Type As MsoConnectorType, BeginX As Single, BeginY As Single, EndX As Single, EndY As Single) As Shape`  
  Creates a connector. Returns a Shape object that represents the new connector. When a connector is added, it's not connected to anything. Use the BeginConnect and EndConnect methods to attach the beginning and end of a connector to other shapes in the document.
    - `Type As MsoConnectorType` (required): The connector type to add.
    - `BeginX As Single` (required): The horizontal position (in points) of the connector's starting point relative to the upper-left corner of the document.
    - `BeginY As Single` (required): The vertical position (in points) of the connector's starting point relative to the upper-left corner of the document.
    - `EndX As Single` (required): The horizontal position (in points) of the connector's end point relative to the upper-left corner of the document.
    - `EndY As Single` (required): The vertical position (in points) of the connector's end point relative to the upper-left corner of the document.
- `AddCurve(SafeArrayOfPoints As Variant) As Shape`  
  Returns a Shape object that represents a Bzier curve on a worksheet.
    - `SafeArrayOfPoints As Variant` (required): An array of coordinate pairs that specifies the vertices and control points of the curve. The first point that you specify is the starting vertex, and the next two points are control points for the first Bzier segment. Then, for each additional segment of the curve, you specify a vertex and two control points. The last point that you specify is the ending vertex for the curve. Note that you must always specify 3n + 1 points, where n is the number of segments in the curve.
- `AddLabel(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Creates a label. Returns a Shape object that represents the new label.
    - `Orientation As MsoTextOrientation` (required): The text orientation within the label.
    - `Left As Single` (required): The position (in points) of the upper-left corner of the label relative to the upper-left corner of the document.
    - `Top As Single` (required): The position (in points) of the upper-left corner of the label relative to the top of the document.
    - `Width As Single` (required): The width of the label, in points.
    - `Height As Single` (required): The height of the label, in points.
- `AddLine(BeginX As Single, BeginY As Single, EndX As Single, EndY As Single) As Shape`  
  As it applies to the Shapes object, returns a Shape object that represents the new line on a worksheet.
    - `BeginX As Single` (required): The position (in points) of the line's starting point relative to the upper-left corner of the document.
    - `BeginY As Single` (required): The position (in points) of the line's starting point relative to the upper-left corner of the document.
    - `EndX As Single` (required): The position (in points) of the line's end point relative to the upper-left corner of the document.
    - `EndY As Single` (required): The position (in points) of the line's end point relative to the upper-left corner of the document.
- `AddPicture(Filename As String, LinkToFile As MsoTriState, SaveWithDocument As MsoTriState, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Creates a picture from an existing file. Returns a Shape object that represents the new picture.
    - `Filename As String` (required): The file from which the picture is to be created.
    - `LinkToFile As MsoTriState` (required): The file to link to. Use msoFalse to make the picture an independent copy of the file. Use msoTrue to link the picture to the file from which it was created.
    - `SaveWithDocument As MsoTriState` (required): To save the picture with the document. Use msoFalse to store only the link information in the document. Use msoTrue to save the linked picture with the document into which it's inserted. This argument must be msoTrue if _LinkToFile_ is msoFalse.
    - `Left As Single` (required): The position (in points) of the upper-left corner of the picture relative to the upper-left corner of the document.
    - `Top As Single` (required): The position (in points) of the upper-left corner of the picture relative to the top of the document.
    - `Width As Single` (required): The width of the picture, in points (enter -1 to retain the width of the existing file).
    - `Height As Single` (required): The height of the picture, in points (enter -1 to retain the height of the existing file).
- `AddPolyline(SafeArrayOfPoints As Variant) As Shape`  
  Creates an open polyline or a closed polygon drawing. Returns a Shape object that represents the new polyline or polygon.
    - `SafeArrayOfPoints As Variant` (required): An array of coordinate pairs that specifies the polyline drawing's vertices.
- `AddShape(Type As MsoAutoShapeType, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Returns a Shape object that represents the new AutoShape on a worksheet.
    - `Type As MsoAutoShapeType` (required): Specifies the type of AutoShape to create.
    - `Left As Single` (required): The position (in points) of the upper-left corner of the AutoShape's bounding box relative to the upper-left corner of the document.
    - `Top As Single` (required): The position (in points) of the upper-left corner of the AutoShape's bounding box relative to the top of the document.
    - `Width As Single` (required): The width of the AutoShape's bounding box, in points.
    - `Height As Single` (required): The height of the AutoShape's bounding box, in points.
- `AddTextEffect(PresetTextEffect As MsoPresetTextEffect, Text As String, FontName As String, FontSize As Single, FontBold As MsoTriState, FontItalic As MsoTriState, Left As Single, Top As Single) As Shape`  
  Creates a WordArt object. Returns a Shape object that represents the new WordArt object.
    - `PresetTextEffect As MsoPresetTextEffect` (required): The preset text effect.
    - `Text As String` (required): The text in the WordArt.
    - `FontName As String` (required): The name of the font used in the WordArt.
    - `FontSize As Single` (required): The size (in points) of the font used in the WordArt.
    - `FontBold As MsoTriState` (required): The font used in the WordArt to bold.
    - `FontItalic As MsoTriState` (required): The font used in the WordArt to italic.
    - `Left As Single` (required): The position (in points) of the upper-left corner of the WordArt's bounding box relative to the upper-left corner of the document.
    - `Top As Single` (required): The position (in points) of the upper-left corner of the WordArt's bounding box relative to the top of the document.
- `AddTextbox(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Creates a text box. Returns a Shape object that represents the new text box.
    - `Orientation As MsoTextOrientation` (required): The orientation of the textbox.
    - `Left As Single` (required): The position (in points) of the upper-left corner of the text box relative to the upper-left corner of the document.
    - `Top As Single` (required): The position (in points) of the upper-left corner of the text box relative to the top of the document.
    - `Width As Single` (required): The width of the text box, in points.
    - `Height As Single` (required): The height of the text box, in points.
- `BuildFreeform(EditingType As MsoEditingType, X1 As Single, Y1 As Single) As FreeformBuilder`  
  Builds a freeform object. Returns a FreeformBuilder object that represents the freeform as it is being built.
    - `EditingType As MsoEditingType` (required): The editing property of the first node.
    - `X1 As Single` (required): The position (in points) of the first node in the freeform drawing relative to the upper-left corner of the document.
    - `Y1 As Single` (required): The position (in points) of the first node in the freeform drawing relative to the upper-left corner of the document.
- `SelectAll()`  
  Selects all the shapes in the specified Shapes collection.
- `AddFormControl(Type As XlFormControl, Left As Long, Top As Long, Width As Long, Height As Long) As Shape`  
  Creates a Microsoft Excel control. Returns a Shape object that represents the new control.
    - `Type As XlFormControl` (required): The Microsoft Excel control type. You cannot create an edit box on a worksheet.
    - `Left As Long` (required): The initial coordinates of the new object (in points) relative to the upper-left corner of cell A1 on a worksheet or to the upper-left corner of a chart.
    - `Top As Long` (required): The initial coordinates of the new object (in points) relative to the top of row 1 on a worksheet, or to the top of the chart area on a chart.
    - `Width As Long` (required): The initial size of the new object, in points.
    - `Height As Long` (required): The initial size of the new object, in points.
- `AddOLEObject([ClassType As Variant], [Filename As Variant], [Link As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As Shape`  
  Creates an OLE object. Returns a Shape object that represents the new OLE object.
    - `ClassType As Variant` (optional): A string that contains the programmatic identifier for the object to be created. You must specify either _ClassType_ or _FileName_. If _ClassType_ is specified, _FileName_ and _Link_ are ignored.
    - `Filename As Variant` (optional): The file from which the object is to be created. If the path isn't specified, the current working folder is used. You must specify either the _ClassType_ or _FileName_ argument for the object, but not both.
    - `Link As Variant` (optional): True to link the OLE object to the file from which it was created. False to make the OLE object an independent copy of the file. If you specified a value for _ClassType_, this argument must be False. The default value is False.
    - `DisplayAsIcon As Variant` (optional): True to display the OLE object as an icon. The default value is False.
    - `IconFileName As Variant` (optional): The file that contains the icon to be displayed.
    - `IconIndex As Variant` (optional): The index of the icon within _IconFileName_. The order of icons in the specified file corresponds to the order in which the icons appear in the Change Icon dialog box (accessed from the Object dialog box when the Display as icon check box is selected). The first icon in the file has the index number 0 (zero). If an icon with the given index number doesn't exist in _IconFileName_, the icon with the index number 1 (the second icon in the file) is used. The default value is 0 (zero).
    - `IconLabel As Variant` (optional): A label (caption) to be displayed beneath the icon.
    - `Left As Variant` (optional): The position (in points) of the upper-left corner of the new object relative to the upper-left corner of the document. The default value is 0 (zero).
    - `Top As Variant` (optional): The position (in points) of the upper-left corner of the new object relative to the top of the document. The default value is 0 (zero).
    - `Width As Variant` (optional): The initial dimensions of the OLE object, in points.
    - `Height As Variant` (optional): The initial dimensions of the OLE object, in points.
- `AddSmartArt(Layout As SmartArtLayout, [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As Shape`  
  Creates a new SmartArt graphic with the specified layout.
    - `Layout As SmartArtLayout` (required): An object that represents the layout to use.
    - `Left As Variant` (optional): The distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).
    - `Top As Variant` (optional): The distance, in points, from the top edge of the object to the top edge of the worksheet.
    - `Width As Variant` (optional): The width, in points, of the object.
    - `Height As Variant` (optional): The height, in points, of the object.
- `AddChart2([Style As Variant], [XlChartType As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], [NewLayout As Variant]) As Shape`  
  Adds a chart to the document. Returns a Shape object that represents a chart and adds it to the specified collection.
    - `Style As Variant` (optional): The chart style. Use "-1" to get the default style for the chart type specified in XlChartType.
    - `XlChartType As Variant` (optional): The type of chart.
    - `Left As Variant` (optional): The position, in points, of the left edge of the chart, relative to the anchor.
    - `Top As Variant` (optional): The position, in points, of the top edge of the chart, relative to the anchor.
    - `Width As Variant` (optional): The width, in points, of the chart.
    - `Height As Variant` (optional): The height, in points, of the chart.
    - `NewLayout As Variant` (optional): If NewLayout is True, the chart is inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).
- `AddPicture2(Filename As String, LinkToFile As MsoTriState, SaveWithDocument As MsoTriState, Left As Single, Top As Single, Width As Single, Height As Single, Compress As MsoPictureCompress) As Shape`  
  Creates a picture from an existing file. Returns a Shape object that represents the new picture.
    - `Filename As String` (required): The file from which the OLE object is to be created.
    - `LinkToFile As MsoTriState` (required): Determines whether the picture will be linked to the file from which it was created.
    - `SaveWithDocument As MsoTriState` (required): Determines whether the linked picture will be saved with the document into which it is inserted. This argument must be msoTrue if _LinkToFile_ is msoFalse.
    - `Left As Single` (required): The position, measured in points, of the left edge of the picture relative to the left edge of the worksheet.
    - `Top As Single` (required): The position, measured in points, of the top edge of the picture relative to the top edge of the worksheet.
    - `Width As Single` (required): The width of the picture, measured in points.
    - `Height As Single` (required): The height of the picture, measured in points.
    - `Compress As MsoPictureCompress` (required): Determines whether the picture should be compressed when inserted.
- `Add3DModel(Filename As String, [LinkToFile As Variant], [SaveWithDocument As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As Shape`  
  Creates a 3D model from an existing file. Returns a Shape object that represents the new 3D model.
    - `Filename As String` (required): The file from which the 3D model is to be created.
    - `LinkToFile As Variant` (optional): Determines whether the 3D model will be linked to the file from which it was created.
    - `SaveWithDocument As Variant` (optional): Determines whether the linked 3D model will be saved with the document into which it is inserted.
    - `Left As Variant` (optional): The position (in points) of the upper-left corner of the 3D model relative to the upper-left corner of the document.
    - `Top As Variant` (optional): The position (in points) of the upper-left corner of the 3D model relative to the top of the document.
    - `Width As Variant` (optional): The width of the 3D model, in points (enter -1 to auto-calculate a width based on the 3D model dimensions).
    - `Height As Variant` (optional): The height of the 3D model, in points (enter -1 to auto-calculate a height based on the 3D model dimensions).
