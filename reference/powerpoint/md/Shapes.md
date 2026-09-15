# Shapes

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493475-5A91-11CF-8700-00AA0060263B}  

A collection of all the Shape objects on the specified slide.

**Remarks:** Each Shape object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes.SelectAll
```

## Properties (8)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `HasTitle As MsoTriState  (read-only)`  
  Returns whether the collection of objects on the specified slide contains a title placeholder. Read-only.
- `Title As Shape  (read-only)`  
  Returns a Shape object that represents the slide title. Read-only.
- `Placeholders As Placeholders  (read-only)`  
  Returns a Placeholders collection that represents the collection of all the placeholders on a slide. Read-only.

## Methods (27)

- `Item(Index As Variant) As Shape`  
  Returns a single Shape object from the specified Shapes collection.
    - `Index As Variant` (required): The name or index number of the single Shape object in the collection to be returned.
- `AddCallout(Type As MsoCalloutType, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Creates a borderless line callout. Returns a Shape object that represents the new callout.
    - `Type As MsoCalloutType` (required): The type of callout line.
    - `Left As Single` (required): The position, measured in points, of the left edge of the callout's bounding box relative to the left edge of the slide.
    - `Top As Single` (required): The position, measured in points, of the top edge of the callout's bounding box relative to the top edge of the slide.
    - `Width As Single` (required): The width of the callout's bounding box, measured in points.
    - `Height As Single` (required): The height of the callout's bounding box, measured in points.
- `AddConnector(Type As MsoConnectorType, BeginX As Single, BeginY As Single, EndX As Single, EndY As Single) As Shape`  
  Creates a connector. Returns a Shape object that represents the new connector. When a connector is added, it is not connected to anything. Use the BeginConnect and EndConnect methods to attach the beginning and end of a connector to other shapes in the document.
    - `Type As MsoConnectorType` (required): The type of connector.
    - `BeginX As Single` (required): The horizontal position, measured in points, of the connector's starting point relative to the left edge of the slide.
    - `BeginY As Single` (required): The vertical position, measured in points, of the connector's starting point relative to the top edge of the slide.
    - `EndX As Single` (required): The horizontal position, measured in points, of the connector's ending point relative to the left edge of the slide.
    - `EndY As Single` (required): The vertical position, measured in points, of the connector's ending point relative to the top edge of the slide.
- `AddCurve(SafeArrayOfPoints As Variant) As Shape`  
  Creates a Bzier curve. Returns a Shape object that represents the new curve.
    - `SafeArrayOfPoints As Variant` (required): An array of coordinate pairs that specifies the vertices and control points of the curve. The first point you specify is the starting vertex, and the next two points are control points for the first Bzier segment. Then, for each additional segment of the curve, you specify a vertex and two control points. The last point you specify is the ending vertex for the curve. Note that you must always specify 3n + 1 points, where n is the number of segments in the curve.
- `AddLabel(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Creates a label. Returns a Shape object that represents the new label.
    - `Orientation As MsoTextOrientation` (required): The text orientation. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `Left As Single` (required): The position, measured in points, of the left edge of the label relative to the left edge of the slide.
    - `Top As Single` (required): The position, measured in points, of the top edge of the label relative to the top edge of the slide.
    - `Width As Single` (required): The width of the label, measured in points.
    - `Height As Single` (required): The height of the label, measured in points.
- `AddLine(BeginX As Single, BeginY As Single, EndX As Single, EndY As Single) As Shape`  
  Creates a line. Returns a Shape object that represents the new line.
    - `BeginX As Single` (required): The horizontal position, measured in points, of the line's starting point relative to the left edge of the slide.
    - `BeginY As Single` (required): The vertical position, measured in points, of the line's starting point relative to the top edge of the slide.
    - `EndX As Single` (required): The horizontal position, measured in points, of the line's ending point relative to the left edge of the slide.
    - `EndY As Single` (required): The vertical position, measured in points, of the line's ending point relative to the top edge of the slide.
- `AddPicture(FileName As String, LinkToFile As MsoTriState, SaveWithDocument As MsoTriState, Left As Single, Top As Single, [Width As Single], [Height As Single]) As Shape`  
  Creates a picture from an existing file. Returns a Shape object that represents the new picture.
    - `FileName As String` (required): The file from which the OLE object is to be created.
    - `LinkToFile As MsoTriState` (required): Determines whether the picture will be linked to the file from which it was created.
    - `SaveWithDocument As MsoTriState` (required): Determines whether the linked picture will be saved with the document into which it is inserted. This argument must be msoTrue if LinkToFile is msoFalse.
    - `Left As Single` (required): The position, measured in points, of the left edge of the picture relative to the left edge of the slide.
    - `Top As Single` (required): The position, measured in points, of the top edge of the picture relative to the top edge of the slide.
    - `Width As Single` (optional): The width of the picture, measured in points.
    - `Height As Single` (optional): The height of the picture, measured in points.
- `AddPolyline(SafeArrayOfPoints As Variant) As Shape`  
  Creates an open polyline or a closed polygon drawing. Returns a Shape object that represents the new polyline or polygon.
    - `SafeArrayOfPoints As Variant` (required): An array of coordinate pairs that specifies the polyline drawing's vertices.
- `AddShape(Type As MsoAutoShapeType, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Creates an AutoShape. Returns a Shape object that represents the new AutoShape.
    - `Type As MsoAutoShapeType` (required): Specifies the type of AutoShape to create.
    - `Left As Single` (required): The position, measured in points, of the left edge of the AutoShape relative to the left edge of the slide.
    - `Top As Single` (required): The position, measured in points, of the top edge of the AutoShape relative to the top edge of the slide.
    - `Width As Single` (required): The width of the AutoShape, measured in points.
    - `Height As Single` (required): The height of the AutoShape, measured in points.
- `AddTextEffect(PresetTextEffect As MsoPresetTextEffect, Text As String, FontName As String, FontSize As Single, FontBold As MsoTriState, FontItalic As MsoTriState, Left As Single, Top As Single) As Shape`  
  Creates a WordArt object. Returns a Shape object that represents the new WordArt object.
    - `PresetTextEffect As MsoPresetTextEffect` (required): The preset text effect.
    - `Text As String` (required): The text in the WordArt.
    - `FontName As String` (required): The name of the font used in the WordArt.
    - `FontSize As Single` (required): The size (in points) of the font used in the WordArt.
    - `FontBold As MsoTriState` (required): Determines whether the font used in the WordArt is set to bold.
    - `FontItalic As MsoTriState` (required): Determines whether the font used in the WordArt is set to italic.
    - `Left As Single` (required): The position, measured in points, of the left edge of the WordArt's bounding box relative to the left edge of the slide.
    - `Top As Single` (required): The position, measured in points, of the top edge of the WordArt's bounding box relative to the top edge of the slide.
- `AddTextbox(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Creates a text box. Returns a Shape object that represents the new text box.
    - `Orientation As MsoTextOrientation` (required): The text orientation. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `Left As Single` (required): The position, measured in points, of the left edge of the text box relative to the left edge of the slide.
    - `Top As Single` (required): The position, measured in points, of the top edge of the text box relative to the top edge of the slide.
    - `Width As Single` (required): The width of the text box, measured in points.
    - `Height As Single` (required): The height of the text box, measured in points.
- `BuildFreeform(EditingType As MsoEditingType, X1 As Single, Y1 As Single) As FreeformBuilder`  
  Builds a freeform object. Returns a FreeformBuilder object that represents the freeform as it is being built.
    - `EditingType As MsoEditingType` (required): The editing property of the first node.
    - `X1 As Single` (required): The horizontal position, measured in points, of the first node in the freeform drawing relative to the left edge of the slide.
    - `Y1 As Single` (required): The vertical position, measured in points, of the first node in the freeform drawing relative to the top edge of the slide.
- `SelectAll()`  
  Selects all the shapes in a Shapes collection.
- `Range([Index As Variant]) As ShapeRange`  
  Returns a ShapeRange object that represents a subset of the shapes in a Shapes collection.
    - `Index As Variant` (optional): The individual shapes that are to be included in the range. Can be an Integer that specifies the index number of the shape, a String that specifies the name of the shape, or an array that contains either integers or strings. If this argument is omitted, the Range method returns all the objects in the specified collection.
- `AddTitle() As Shape`  
  Restores a previously deleted title placeholder to a slide. Returns a Shape object that represents the restored title.
- `AddOLEObject([Left As Single], [Top As Single], [Width As Single], [Height As Single], [ClassName As String], [FileName As String], [DisplayAsIcon As MsoTriState], [IconFileName As String], [IconIndex As Long], [IconLabel As String], [Link As MsoTriState]) As Shape`  
  Creates an OLE object. Returns a Shape object that represents the new OLE object.
    - `Left As Single` (optional): The position (in points) of the upper-left corner of the new object relative to the upper-left corner of the slide. The default value is 0 (zero).
    - `Top As Single` (optional): The position (in points) of the upper-left corner of the new object relative to the upper-left corner of the slide. The default value is 0 (zero).
    - `Width As Single` (optional): The initial width of the OLE object, in points.
    - `Height As Single` (optional): The initial height of the OLE object, in points.
    - `ClassName As String` (optional): The OLE long class name or the ProgID for the object that's to be created. You must specify either the ClassName or FileName argument for the object, but not both.
    - `FileName As String` (optional): The file from which the object is to be created. If the path isn't specified, the current working folder is used. You must specify either the ClassName or FileName argument for the object, but not both.
    - `DisplayAsIcon As MsoTriState` (optional): Determines whether the OLE object will be displayed as an icon.
    - `IconFileName As String` (optional): The file that contains the icon to be displayed.
    - `IconIndex As Long` (optional): The index of the icon within IconFileName. The first icon in the file has the index number 0 (zero). If an icon with the given index number doesn't exist in IconFileName, the icon with the index number 1 (the second icon in the file) is used. The default value is 0 (zero).
    - `IconLabel As String` (optional): A label (caption) to be displayed beneath the icon.
    - `Link As MsoTriState` (optional): Determines whether the OLE object will be linked to the file from which it was created. If you specified a value for ClassName, this argument must be msoFalse.
- `AddPlaceholder(Type As PpPlaceholderType, [Left As Single], [Top As Single], [Width As Single], [Height As Single]) As Shape`  
  Restores a previously deleted placeholder on a slide. Returns a Shape object that represents the restored placeholder.
    - `Type As PpPlaceholderType` (required): The type of placeholder. Placeholders of type ppPlaceholderVerticalBody or ppPlaceholderVerticalTitle are found only on slides of layout type ppLayoutVerticalText, ppLayoutClipArtAndVerticalText, ppLayoutVerticalTitleAndText, or ppLayoutVerticalTitleAndTextOverChart. You cannot create slides with any of these layouts from the user interface; you must create them programmatically by using the Add method or by setting the Layout property of an existing slide.
    - `Left As Single` (optional): The position (in points) of the upper-left corner of the placeholder relative to the upper-left corner of the document.
    - `Top As Single` (optional): The position (in points) of the upper-left corner of the placeholder relative to the upper-left corner of the document.
    - `Width As Single` (optional): The width of the placeholder, in points.
    - `Height As Single` (optional): The height of the placeholder, in points.
- `Paste() As ShapeRange`  
  Pastes the shapes, slides, or text on the Clipboard into the specified Shapes collection, at the top of the z-order. Each pasted object becomes a member of the specified Shapes collection. If the Clipboard contains entire slides, the slides will be pasted as shapes that contain the images of the slides. If the Clipboard contains a text range, the text will be pasted into a newly created TextFrame shape. Returns a ShapeRange object that represents the pasted objects.
- `AddTable(NumRows As Long, NumColumns As Long, [Left As Single], [Top As Single], [Width As Single], [Height As Single]) As Shape`  
  Adds a table shape to a slide.
    - `NumRows As Long` (required): The number of rows in the table.
    - `NumColumns As Long` (required): The number of columns in the table.
    - `Left As Single` (optional): The distance (in points) from the left edge of the slide to the left edge of the table.
    - `Top As Single` (optional): The distance (in points) from the top edge of the slide to the top edge of the table.
    - `Width As Single` (optional): The width (in points) of the new table.
    - `Height As Single` (optional): The height (in points) of the new table.
- `PasteSpecial([DataType As PpPasteDataType], [DisplayAsIcon As MsoTriState], [IconFileName As String], [IconIndex As Long], [IconLabel As String], [Link As MsoTriState]) As ShapeRange`  
  Pastes the contents of the Clipboard, using a special format.
    - `DataType As PpPasteDataType` (optional): A format for the Clipboard contents when they're inserted into the document. The default value varies, depending on the contents in the Clipboard. An error occurs if the specified data type in the DataType argument is not supported by the clipboard contents.
    - `DisplayAsIcon As MsoTriState` (optional): MsoTrue to display the embedded object (or link) as an icon.
    - `IconFileName As String` (optional): If DisplayAsIcon is set to msoTrue, this argument is the path and file name for the file in which the icon to be displayed is stored. If DisplayAsIcon is set to msoFalse, this argument is ignored.
    - `IconIndex As Long` (optional): If DisplayAsIcon is set to msoTrue, this argument is a number that corresponds to the icon you want to use in the program file specified by IconFilename. For example, 0 (zero) corresponds to the first icon, and 1 corresponds to the second icon. If this argument is omitted, the first (default) icon is used. If DisplayAsIcon is set to msoFalse, then this argument is ignored. If IconIndex is outside the valid range, then the default icon (index 0) is used.
    - `IconLabel As String` (optional): If DisplayAsIcon is set to msoTrue, this argument is the text that appears below the icon. If this label is missing, Microsoft PowerPoint generates an icon label based on the Clipboard contents. If DisplayAsIcon is set to msoFalse, then this argument is ignored.
    - `Link As MsoTriState` (optional): Determines whether to create a link to the source file of the Clipboard contents. An error occurs if the Clipboard contents don't support a link.
- `AddMediaObject2(FileName As String, [LinkToFile As MsoTriState], [SaveWithDocument As MsoTriState], [Left As Single], [Top As Single], [Width As Single], [Height As Single]) As Shape`  
  Replaces deprecated Shapes.AddMediaObject method (PowerPoint). Adds a new media object.
    - `FileName As String` (required): The name of the file to be added.
    - `LinkToFile As MsoTriState` (optional): Indicates whether to link to the file.
    - `SaveWithDocument As MsoTriState` (optional): Indicates whether to save the media with the document.
    - `Left As Single` (optional): The distance, in points, from the left edge of the slide to the left edge of the media object.
    - `Top As Single` (optional): The distance, in points, from the top edge of the slide to the top edge of the media object.
    - `Width As Single` (optional): The width, in points, of the media object. Default value is -1.
    - `Height As Single` (optional): The height, in points, of the media object. Default value is -1.
- `AddMediaObjectFromEmbedTag(EmbedTag As String, [Left As Single], [Top As Single], [Width As Single], [Height As Single]) As Shape`  
  Adds a media object from an embedded tag to a Shapes object.
    - `EmbedTag As String` (required): The embed tag.
    - `Left As Single` (optional): The distance, in points, from the left edge of the slide to the left edge of the media object.
    - `Top As Single` (optional): The distance, in points, from the top edge of the slide to the top edge of the media object.
    - `Width As Single` (optional): The width, in points, of the media object.
    - `Height As Single` (optional): The height, in points, of the media object.
- `AddSmartArt(Layout As SmartArtLayout, [Left As Single], [Top As Single], [Width As Single], [Height As Single]) As Shape`  
  Adds a SmartArt diagram to the Shapes object.
    - `Layout As SmartArtLayout` (required): The SmartArt diagram to add.
    - `Left As Single` (optional): The distance, in points, from the left edge of the slide to the left edge of the SmartArt diagram.
    - `Top As Single` (optional): The distance, in points, from the top edge of the slide to the top edge of the SmartArt diagram.
    - `Width As Single` (optional): The width of the SmartArt diagram.
    - `Height As Single` (optional): The height of the SmartArt diagram.
- `AddChart2([Style As Long], [Type As XlChartType], [Left As Single], [Top As Single], [Width As Single], [Height As Single], [NewLayout As Boolean]) As Shape`  
  Adds a chart to the document. Returns a Shape object that represents a chart and adds it to the specified collection.
- `AddPicture2(FileName As String, LinkToFile As MsoTriState, SaveWithDocument As MsoTriState, Left As Single, Top As Single, [Width As Single], [Height As Single], [compress As MsoPictureCompress]) As Shape`  
  Creates a picture from an existing file. Returns a Shape object that represents the new picture.
    - `FileName As String` (required): The file from which the OLE object is to be created.
    - `LinkToFile As MsoTriState` (required): Determines whether the picture will be linked to the file from which it was created.
    - `SaveWithDocument As MsoTriState` (required): Determines whether the linked picture will be saved with the document into which it is inserted. This argument must be msoTrue if _LinkToFile_ is msoFalse.
    - `Left As Single` (required): The position, measured in points, of the left edge of the picture relative to the left edge of the slide.
    - `Top As Single` (required): The position, measured in points, of the top edge of the picture relative to the top edge of the slide.
    - `Width As Single` (optional): The width of the picture, measured in points.
    - `Height As Single` (optional): The height of the picture, measured in points.
    - `compress As MsoPictureCompress` (optional): Determines whether the picture should be compressed when inserted.
- `AddInkShapeFromXML(InkXML As String, Left As Single, Top As Single, [Width As Single], [Height As Single]) As Shape`  
  Creates an ink shape. Returns a Shape object that represents the new ink shape.
    - `Left As Single` (required): The position, measured in points, of the left edge of the ink shape relative to the left edge of the slide.
    - `Top As Single` (required): The position, measured in points, of the top edge of the ink shape relative to the top edge of the slide.
    - `Width As Single` (optional): The width of the ink shape, measured in points. If this parameter is not specified, the width is calculated based off of the InkActionML.
    - `Height As Single` (optional): The height of the ink shape, measured in points. If this parameter is not specified, the hight is calculated based off of the InkActionML.
- `Add3DModel(FileName As String, LinkToFile As MsoTriState, SaveWithDocument As MsoTriState, Left As Single, Top As Single, [Width As Single], [Height As Single]) As Shape`  
  Creates a Model3DFormat object from an existing file. Returns a Shape object that represents the new 3D model.
    - `FileName As String` (required): The file from which the 3D model object is to be created.
    - `LinkToFile As MsoTriState` (required): Determines whether the 3D model will be linked to the file from which it was created.
    - `SaveWithDocument As MsoTriState` (required): Determines whether the linked 3D model will be saved with the document into which it is inserted. This argument must be msoTrue if LinkToFile is msoFalse.
    - `Left As Single` (required): The position, measured in points, of the left edge of the 3D model relative to the left edge of the slide.
    - `Top As Single` (required): The position, measured in points, of the top edge of the 3D model relative to the top edge of the slide.
    - `Width As Single` (optional): The width of the 3D model, measured in points (enter -1 to auto-calculate a width based on the 3D model dimensions).
    - `Height As Single` (optional): The height of the 3D model, measured in points (enter -1 to auto-calculate a height based on the 3D model dimensions).
