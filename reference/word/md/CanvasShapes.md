# CanvasShapes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {396F9073-F9FD-11D3-8EA0-0050049A1A01}  

Use the CanvasItems property of either a Shape or ShapeRange object to return a CanvasShapes collection.

**Remarks:** To add shapes to a drawing canvas, use the following methods of the CanvasShapes collection: AddCallout, AddConnector, AddCurve, AddLabel, AddLine, AddPicture, AddPolyline, AddShape, AddTextbox, AddTextEffect, or BuildFreeform. The following example adds a drawing canvas to the active document and then adds three shapes to the drawing canvas. Use CanvasItems (index), where index is the name or the index number, to return a single shape in the CanvasShapes collection. The following example sets the Line and Fill properties and vertically flips the third shape in a drawing canvas. Each shape is assigned a default name when it is created. For example, if you add three different shapes to a document, they might be named Rectangle 2, TextBox 3, and Oval 4. Use the Name property to reference the default name or to assign a more meaningful name to a shape.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CanvasShapes object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of canvas shapes in the specified collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (14)

- `Item(Index As Variant) As Shape`  
  Returns an individual Shape object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `AddCallout(Type As MsoCalloutType, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Adds a borderless line callout to a drawing canvas. Returns a Shape object that represents the callout.
    - `Type As MsoCalloutType` (required): The type of callout.
    - `Left As Single` (required): The position, in points, of the left edge of the callout's bounding box.
    - `Top As Single` (required): The position, in points, of the top edge of the callout's bounding box.
    - `Width As Single` (required): The width, in points, of the callout's bounding box.
    - `Height As Single` (required): The height, in points, of the callout's bounding box.
- `AddConnector(Type As MsoConnectorType, BeginX As Single, BeginY As Single, EndX As Single, EndY As Single) As Shape`  
  Returns a Shape object that represents a connecting line between two shapes in a drawing canvas.
    - `Type As MsoConnectorType` (required): The type of connector.
    - `BeginX As Single` (required): The horizontal position that marks the beginning of the connector.
    - `BeginY As Single` (required): The vertical position that marks the beginning of the connector.
    - `EndX As Single` (required): The horizontal position that marks the end of the connector.
    - `EndY As Single` (required): The vertical position that marks the end of the connector.
- `AddCurve(SafeArrayOfPoints As Variant) As Shape`  
  Returns a Shape object that represents a Bzier curve in a drawing canvas.
    - `SafeArrayOfPoints As Variant` (required): An array of coordinate pairs that specifies the vertices and control points of the curve. The first point you specify is the starting vertex, and the next two points are control points for the first Bzier segment. Then, for each additional segment of the curve, you specify a vertex and two control points. The last point you specify is the ending vertex for the curve. Note that you must always specify 3n + 1 points, where n is the number of segments in the curve.
- `AddLabel(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Adds a text label to a drawing canvas. Returns a Shapes object that represents the text label.
    - `Orientation As MsoTextOrientation` (required): The orientation of the text.
    - `Left As Single` (required): The position, measured in points, of the left edge of the label relative to the left edge of the drawing canvas.
    - `Top As Single` (required): The position, measured in points, of the top edge of the label relative to the top edge of the drawing canvas.
    - `Width As Single` (required): The width of the label, in points.
    - `Height As Single` (required): The height of the label, in points.
- `AddLine(BeginX As Single, BeginY As Single, EndX As Single, EndY As Single) As Shape`  
  Adds a line to a drawing canvas. Returns a Shape object that represents the line and adds it to the CanvasShapes collection.
    - `BeginX As Single` (required): The horizontal position, measured in points, of the line's starting point, relative to the drawing canvas.
    - `BeginY As Single` (required): The vertical position, measured in points, of the line's starting point, relative to the drawing canvas.
    - `EndX As Single` (required): The horizontal position, measured in points, of the line's endpoint, relative to the drawing canvas.
    - `EndY As Single` (required): The vertical position, measured in points, of the line's endpoint, relative to the drawing canvas.
- `AddPicture(FileName As String, [LinkToFile As Variant], [SaveWithDocument As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As Shape`  
  Adds a picture to a drawing canvas. Returns a Shape object that represents the picture and adds it to the CanvasShapes collection.
    - `FileName As String` (required): The path and file name of the picture.
    - `LinkToFile As Variant` (optional): True to link the picture to the file from which it was created. False to make the picture an independent copy of the file. The default value is False.
    - `SaveWithDocument As Variant` (optional): True to save the linked picture with the document. The default value is False.
    - `Left As Variant` (optional): The position, measured in points, of the left edge of the new picture relative to the drawing canvas.
    - `Top As Variant` (optional): The position, measured in points, of the top edge of the new picture relative to the drawing canvas.
    - `Width As Variant` (optional): The width of the picture, in points.
    - `Height As Variant` (optional): The height of the picture, in points.
- `AddPolyline(SafeArrayOfPoints As Variant) As Shape`  
  Adds an open or closed polygon to a drawing canvas. Returns a Shape object that represents the polygon.
    - `SafeArrayOfPoints As Variant` (required): An array of coordinate pairs that specifies the polyline drawing's vertices.
- `AddShape(Type As Long, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Adds an AutoShape to a drawing canvas. Returns a Shape object that represents the AutoShape.
    - `Type As Long` (required): The type of shape to be returned. Can be any MsoAutoShape constant.
    - `Left As Single` (required): The position, measured in points, of the left edge of the AutoShape.
    - `Top As Single` (required): The position, measured in points, of the top edge of the AutoShape.
    - `Width As Single` (required): The width, measured in points, of the AutoShape.
    - `Height As Single` (required): The height, measured in points, of the AutoShape.
- `AddTextEffect(PresetTextEffect As MsoPresetTextEffect, Text As String, FontName As String, FontSize As Single, FontBold As MsoTriState, FontItalic As MsoTriState, Left As Single, Top As Single) As Shape`  
  Adds a WordArt shape to a drawing canvas. Returns a Shape object that represents the WordArt.
    - `PresetTextEffect As MsoPresetTextEffect` (required): A preset text effect. The values of the MsoPresetTextEffect constants correspond to the formats listed in the WordArt Gallery dialog box (numbered from left to right and from top to bottom).
    - `Text As String` (required): The text in the WordArt.
    - `FontName As String` (required): The name of the font used in the WordArt.
    - `FontSize As Single` (required): The size (in points) of the font used in the WordArt.
    - `FontBold As MsoTriState` (required): msoTrue to bold the WordArt font. msoFalse to set the font used in the WordArt to regular.
    - `FontItalic As MsoTriState` (required): msoTrue to italicize the WordArt font. msoFalse to set the font used in the WordArt to regular.
    - `Left As Single` (required): The position, measured in points, of the left edge of the WordArt shape relative to the left edge of the drawing canvas.
    - `Top As Single` (required): The position, measured in points, of the top edge of the WordArt shape relative to the top edge of the drawing canvas.
- `AddTextbox(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single) As Shape`  
  Adds a text box to a drawing canvas. Returns a Shape object that represents the text box.
    - `Orientation As MsoTextOrientation` (required): The orientation of the text. Some of the MsoTextOrientation constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `Left As Single` (required): The position, measured in points, of the left edge of the text box.
    - `Top As Single` (required): The position, measured in points, of the top edge of the text box.
    - `Width As Single` (required): The width, measured in points, of the text box.
    - `Height As Single` (required): The height, measured in points, of the text box.
- `BuildFreeform(EditingType As MsoEditingType, X1 As Single, Y1 As Single) As FreeformBuilder`  
  Builds a freeform object. Returns a FreeformBuilder object that represents the freeform as it is being built. .
    - `EditingType As MsoEditingType` (required): The EditingType parameter can be either msoEditingAuto or msoEditingCorner; cannot be msoEditingSmooth or msoEditingSymmetric.
    - `X1 As Single` (required): The position (in points) of the first node in the freeform drawing relative to the left edge of the document.
    - `Y1 As Single` (required): The position (in points) of the first node in the freeform drawing relative to the top of the document.
- `Range(Index As Variant) As ShapeRange`  
  Returns a ShapeRange object.
    - `Index As Variant` (required): Specifies which shapes are to be included in the specified range. Can be an integer that specifies the index number of a shape within the Shapes collection, a string that specifies the name of a shape, or a array that contains integers or strings.
- `SelectAll()`  
  Selects all the shapes in a canvas.
