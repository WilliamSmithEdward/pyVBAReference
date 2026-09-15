# Shapes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002099F-0000-0000-C000-000000000046}  

A collection of Shape objects that represent all the shapes in a document or all the shapes in all the headers and footers in a document. Each Shape object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.

**Remarks:** If you want to work with a subset of the shapes on a document - for example, to do something to only the AutoShapes on the document or to only the selected shapes - you must construct a ShapeRange collection that contains the shapes you want to work with. Use the Shapes property to return the Shapes collection. The following example selects all the shapes on the active document. Use one of the following methods of the Shapes collection: Add3DModel, AddCallout, AddCurve, AddLabel, AddLine, AddOleControl, AddOleObject, AddPolyline, AddShape, AddTextbox, AddTextEffect, or BuildFreeForm to add a shape to a document return a Shape object that represents the newly created shape. The following example adds a rectangle to the active document. Use Shapes (Index), where Index is the name or the index number, to return a single Shape object. The following example horizontally flips shape one on the active document. This example horizontally flips the shape named "Rectangle 1" on the active document. Each shape is assigned a default name when it is created. For example, if you add three different shapes to a document, they might be named "Rectangle 2," "TextBox 3," and "Oval 4.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Shapes object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of shapes in the collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (20)

- `Item(Index As Variant) As Shape`  
  Returns an individual Shape object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `AddCallout(Type As MsoCalloutType, Left As Single, Top As Single, Width As Single, Height As Single, [Anchor As Variant]) As Shape`  
  Adds a borderless line callout to a drawing canvas.
    - `Type As MsoCalloutType` (required): The type of callout.
    - `Left As Single` (required): The position, in points, of the left edge of the callout's bounding box.
    - `Top As Single` (required): The position, in points, of the top edge of the callout's bounding box.
    - `Width As Single` (required): The width, in points, of the callout's bounding box.
    - `Height As Single` (required): The height, in points, of the callout's bounding box.
- `AddCurve(SafeArrayOfPoints As Variant, [Anchor As Variant]) As Shape`  
  Returns a Shape object that represents a Bzier curve in a drawing canvas.
    - `SafeArrayOfPoints As Variant` (required): An array of coordinate pairs that specifies the vertices and control points of the curve. The first point you specify is the starting vertex, and the next two points are control points for the first Bzier segment. Then, for each additional segment of the curve, you specify a vertex and two control points. The last point you specify is the ending vertex for the curve. Note that you must always specify 3n + 1 points, where n is the number of segments in the curve.
- `AddLabel(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single, [Anchor As Variant]) As Shape`  
  Adds a text label to a drawing canvas.
    - `Orientation As MsoTextOrientation` (required): The orientation of the text.
    - `Left As Single` (required): The position, measured in points, of the left edge of the label relative to the left edge of the drawing canvas.
    - `Top As Single` (required): The position, measured in points, of the top edge of the label relative to the top edge of the drawing canvas.
    - `Width As Single` (required): The width of the label, in points.
    - `Height As Single` (required): The height of the label, in points.
- `AddLine(BeginX As Single, BeginY As Single, EndX As Single, EndY As Single, [Anchor As Variant]) As Shape`  
  Adds a line to a drawing canvas.
    - `BeginX As Single` (required): The horizontal position, measured in points, of the line's starting point, relative to the drawing canvas.
    - `BeginY As Single` (required): The vertical position, measured in points, of the line's starting point, relative to the drawing canvas.
    - `EndX As Single` (required): The horizontal position, measured in points, of the line's endpoint, relative to the drawing canvas.
    - `EndY As Single` (required): The vertical position, measured in points, of the line's endpoint, relative to the drawing canvas.
- `AddPicture(FileName As String, [LinkToFile As Variant], [SaveWithDocument As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], [Anchor As Variant]) As Shape`  
  Adds a picture to a document. Returns a Shape object that represents the picture and adds it to the Shapes collection.
    - `FileName As String` (required): The path and file name of the picture.
    - `LinkToFile As Variant` (optional): True to link the picture to the file from which it was created. False to make the picture an independent copy of the file. The default value is False.
    - `SaveWithDocument As Variant` (optional): True to save the linked picture with the document. The default value is False.
    - `Left As Variant` (optional): The position, measured in points, of the left edge of the new picture relative to the drawing canvas.
    - `Top As Variant` (optional): The position, measured in points, of the top edge of the new picture relative to the drawing canvas.
    - `Width As Variant` (optional): The width of the picture, in points.
    - `Height As Variant` (optional): The height of the picture, in points.
    - `Anchor As Variant` (optional): A Range in the document.
- `AddPolyline(SafeArrayOfPoints As Variant, [Anchor As Variant]) As Shape`  
  Adds an open or closed polygon to a drawing canvas. Returns a Shape object that represents the polygon and adds it to the CanvasShapes collection.
    - `SafeArrayOfPoints As Variant` (required): An array of coordinate pairs that specifies the polyline drawing's vertices.
- `AddShape(Type As Long, Left As Single, Top As Single, Width As Single, Height As Single, [Anchor As Variant]) As Shape`  
  Adds an AutoShape to a document. Returns a Shape object that represents the AutoShape and adds it to the Shapes collection.
    - `Type As Long` (required): The type of shape to be returned. Can be any MsoAutoShapeType constant.
    - `Left As Single` (required): The position, measured in points, of the left edge of the AutoShape.
    - `Top As Single` (required): The position, measured in points, of the top edge of the AutoShape.
    - `Width As Single` (required): The width, measured in points, of the AutoShape.
    - `Height As Single` (required): The height, measured in points, of the AutoShape.
- `AddTextEffect(PresetTextEffect As MsoPresetTextEffect, Text As String, FontName As String, FontSize As Single, FontBold As MsoTriState, FontItalic As MsoTriState, Left As Single, Top As Single, [Anchor As Variant]) As Shape`  
  Adds a WordArt shape to a drawing canvas. Returns a Shape object that represents the WordArt and adds it to the CanvasShapes collection.
    - `PresetTextEffect As MsoPresetTextEffect` (required): A preset text effect. The values of the MsoPresetTextEffect constants correspond to the formats listed in the WordArt Gallery dialog box (numbered from left to right and from top to bottom).
    - `Text As String` (required): The text in the WordArt.
    - `FontName As String` (required): The name of the font used in the WordArt.
    - `FontSize As Single` (required): The size (in points) of the font used in the WordArt.
    - `FontBold As MsoTriState` (required): MsoTrue to bold the WordArt font.
    - `FontItalic As MsoTriState` (required): MsoTrue to italicize the WordArt font.
    - `Left As Single` (required): The position, measured in points, of the left edge of the WordArt shape relative to the left edge of the drawing canvas.
    - `Top As Single` (required): The position, measured in points, of the top edge of the WordArt shape relative to the top edge of the drawing canvas.
- `AddTextbox(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single, [Anchor As Variant]) As Shape`  
  Adds a text box to a drawing canvas.
    - `Orientation As MsoTextOrientation` (required): The orientation of the text. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `Left As Single` (required): The position, measured in points, of the left edge of the text box.
    - `Top As Single` (required): The position, measured in points, of the top edge of the text box.
    - `Width As Single` (required): The width, measured in points, of the text box.
    - `Height As Single` (required): The height, measured in points, of the text box.
- `BuildFreeform(EditingType As MsoEditingType, X1 As Single, Y1 As Single) As FreeformBuilder`  
  Builds a freeform object.
    - `EditingType As MsoEditingType` (required): The editing property of the first node.
    - `X1 As Single` (required): The position (in points) of the first node in the freeform drawing relative to the left edge of the document.
    - `Y1 As Single` (required): The position (in points) of the first node in the freeform drawing relative to the top edge of the document.
- `Range(Index As Variant) As ShapeRange`  
  Returns a ShapeRange object that represents the shapes within a range.
    - `Index As Variant` (required): Specifies which shapes are to be included in the specified range. Can be an integer that specifies the index number of a shape within the Shapes collection, a string that specifies the name of a shape, or a array that contains integers or strings.
- `SelectAll()`  
  Selects all the shapes in a collection of shapes.
- `AddOLEObject([ClassType As Variant], [FileName As Variant], [LinkToFile As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], [Anchor As Variant]) As Shape`  
  Creates an OLE object. Returns the InlineShape object that represents the new OLE object.
    - `ClassType As Variant` (optional): The name of the application used to activate the specified OLE object.
    - `FileName As Variant` (optional): The file from which the object is to be created. If this argument is omitted, the current folder is used. You must specify either the ClassType or FileName argument for the object, but not both.
    - `LinkToFile As Variant` (optional): True to link the OLE object to the file from which it was created. False to make the OLE object an independent copy of the file. If you specified a value for ClassType, the LinkToFile argument must be False. The default value is False.
    - `DisplayAsIcon As Variant` (optional): True to display the OLE object as an icon. The default value is False.
    - `IconFileName As Variant` (optional): The file that contains the icon to be displayed.
    - `IconIndex As Variant` (optional): The index number of the icon within IconFileName. The order of icons in the specified file corresponds to the order in which the icons appear in the Change Icon dialog box when the Display as icon check box is selected. The first icon in the file has the index number 0 (zero). If an icon with the given index number doesn't exist in IconFileName, the icon with the index number 1 (the second icon in the file) is used. The default value is 0 (zero).
    - `IconLabel As Variant` (optional): A label (caption) to be displayed beneath the icon.
- `AddOLEControl([ClassType As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], [Anchor As Variant]) As Shape`  
  Creates an ActiveX control (formerly known as an OLE control). Returns the InlineShape object that represents the new ActiveX control.
    - `ClassType As Variant` (optional): The programmatic identifier for the ActiveX control to be created.
- `AddCanvas(Left As Single, Top As Single, Width As Single, Height As Single, [Anchor As Variant]) As Shape`  
  Adds a drawing canvas to a document. Returns a Shape object that represents the drawing canvas and adds it to the Shapes collection.
    - `Left As Single` (required): The position, in points, of the left edge of the drawing canvas, relative to the anchor.
    - `Top As Single` (required): The position, in points, of the top edge of the drawing canvas, relative to the anchor.
    - `Width As Single` (required): The width, in points, of the drawing canvas.
    - `Height As Single` (required): The height, in points, of the drawing canvas.
    - `Anchor As Variant` (optional): A Range object that represents the text to which the canvas is bound. If Anchor is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the canvas is positioned relative to the top and left edges of the page.
- `AddSmartArt(Layout As SmartArtLayout, [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], [Anchor As Variant]) As Shape`  
  Inserts the specified SmartArt graphic into the active document.
    - `Layout As SmartArtLayout` (required): A SmartArtLayout object that specifies the layout for the SmartArt graphic.
    - `Left As Variant` (optional): The distance, in points, from the left edge of the slide to the left edge of the SmartArt graphic.
    - `Top As Variant` (optional): The distance, in points, from the top edge of the slide to the top edge of the SmartArt graphic.
    - `Width As Variant` (optional): The width of the SmartArt graphic.
    - `Height As Variant` (optional): The height of the SmartArt graphic.
    - `Anchor As Variant` (optional): A Range object that represents the text to which the SmartArt graphic is bound. If Anchor is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the SmartArt graphic is positioned relative to the top and left edges of the page.
- `AddWebVideo(EmbedCode As String, VideoWidth As Variant, VideoHeight As Variant, [PosterFrameImage As Variant], [Url As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], [Anchor As Variant]) As Shape`  
  Adds a new web video to the document.
    - `EmbedCode As String` (required): The HTML code to embed.
    - `VideoWidth As Variant` (required): An integer that represents the width of the web video in pixels.
    - `VideoHeight As Variant` (required): An integer that represents the height of the web video in pixels.
    - `PosterFrameImage As Variant` (optional): A string that points to the file to use as the poster frame for the web video.
    - `Url As Variant` (optional): A string that contains the URL to the web video.
    - `Left As Variant` (optional): The position, measured in points, of the left edge of the poster frame from the edge of the document.
    - `Top As Variant` (optional): The position, measured in points, of the top edge of the poster frame from the edge of the document.
    - `Width As Variant` (optional): The width, measured in points, of the poster frame in the document.
    - `Height As Variant` (optional): The height, measured in points, of the poster frame in the document.
    - `Anchor As Variant` (optional): A Range object that represents the text to which the web video is bound. If Anchor is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the video is positioned relative to the top and left edges of the page.
- `AddChart2([Style As Long], [Type As XlChartType], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], [Anchor As Variant], [NewLayout As Variant]) As Shape`  
  Adds a chart to the document. Returns a Shape object that represents a chart and adds it to the specified collection.
    - `Style As Long` (optional): The chart style. Use "-1" to get the default style for the chart type specified in Type.
    - `Type As XlChartType` (optional): The type of chart.
    - `Left As Variant` (optional): The position, in points, of the left edge of the chart, relative to the anchor.
    - `Top As Variant` (optional): The position, in points, of the top edge of the chart, relative to the anchor.
    - `Width As Variant` (optional): The width, in points, of the chart.
    - `Height As Variant` (optional): The height, in points, of the chart.
    - `Anchor As Variant` (optional): A Range object that represents the text to which the canvas is bound. If Anchor is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the canvas is positioned relative to the top and left edges of the page.
    - `NewLayout As Variant` (optional): If NewLayout is True, the chart will be inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).
- `Add3DModel(FileName As String, [LinkToFile As Variant], [SaveWithDocument As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant], [Anchor As Variant]) As Shape`  
  Adds a 3D model to a drawing canvas. Returns a Shape object that represents the 3D model and adds it to the CanvasShapes collection.
    - `FileName As String` (required): The path and file name of the 3D model.
    - `LinkToFile As Variant` (optional): True to link the 3D model to the file from which it was created. False to make the 3D model an independent copy of the file. The default value is False.
    - `SaveWithDocument As Variant` (optional): True to save the linked 3D model with the document. The default value is False.
    - `Left As Variant` (optional): The position, measured in points, of the left edge of the new 3D model relative to the drawing canvas.
    - `Top As Variant` (optional): The position, measured in points, of the top edge of the new 3D model relative to the drawing canvas.
    - `Width As Variant` (optional): The width of the 3D model, in points (enter -1 to auto-calculate a width based on the 3D model dimensions).
    - `Height As Variant` (optional): The height of the 3D model, in points (enter -1 to auto-calculate a height based on the 3D model dimensions).
