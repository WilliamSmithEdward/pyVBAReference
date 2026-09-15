# Shape

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209A0-0000-0000-C000-000000000046}  

Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, ActiveX control, or picture. The Shape object is a member of the Shapes collection, which includes all the shapes in the main story of a document or in all the headers and footers of a document.

**Remarks:** A shape is always attached to an anchoring range. You can position the shape anywhere on the page that contains the anchor. There are three objects that represent shapes: the Shapes collection, which represents all the shapes on a document; the ShapeRange object, which represents a specified subset of the shapes on a document (for example, a ShapeRange object could represent shapes one and four on the document, or it could represent all the selected shapes on the document); and the Shape object, which represents a single shape on a document. If you want to work with several shapes at the same time or with shapes within the selection, use a ShapeRange collection. Use Shapes (index), where index is the name or the index number, to return a single Shape object. The following example horizontally flips shape one on the active document. The following example horizontally flips the shape named Rectangle 1 on the active document. Each shape is assigned a default name when it is created. For example, if you add three different shapes to a document, they might be named Rectangle 2, TextBox 3, and Oval 4. To give a shape a more meaningful name, set the Name property.

## Properties (63)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Shape object.
- `Adjustments As Adjustments  (read-only)`  
  Returns an Adjustments object that contains adjustment values for all the adjustments in the specified Shape object that represents an AutoShape or WordArt. Read-only.
- `AutoShapeType As MsoAutoShapeType  (read/write)`  
  Returns or sets the shape type for the specified Shape object, which must represent an AutoShape other than a line or freeform drawing. Read/write MsoAutoShapeType.
- `Callout As CalloutFormat  (read-only)`  
  Returns a CalloutFormat object that contains callout formatting properties for the specified shape. Read-only.
- `Fill As FillFormat  (read-only)`  
  Returns a FillFormat object that contains fill formatting properties for the specified shape. Read-only.
- `GroupItems As GroupShapes  (read-only)`  
  Returns a GroupShapes object that represents the individual shapes in the specified group. Read-only.
- `Height As Single  (read/write)`  
  Returns or sets the height (in points) of the specified shape. Read/write Single.
- `HorizontalFlip As MsoTriState  (read-only)`  
  Indicates that a shape has been flipped horizontally. Read-only MsoTriState.
- `Left As Single  (read/write)`  
  Returns or sets a Single that represents the horizontal position, measured in points, of the specified shape or shape range. Can also be any valid WdShapePosition constant. Read/write.
- `Line As LineFormat  (read-only)`  
  Returns a LineFormat object that contains line formatting properties for the specified shape. Read-only.
- `LockAspectRatio As MsoTriState  (read/write)`  
  MsoTrue if the specified shape retains its original proportions when you resize it. MsoFalse if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write String.
- `Nodes As ShapeNodes  (read-only)`  
  Returns a ShapeNodes collection that represents the geometric description of the specified shape.
- `Rotation As Single  (read/write)`  
  Returns or sets the number of degrees the specified shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation. Read/write Single.
- `PictureFormat As PictureFormat  (read-only)`  
  Returns a PictureFormat object that contains picture formatting properties for the specified object. Read-only.
- `Shadow As ShadowFormat  (read-only)`  
  Returns a ShadowFormat object that represents the shadow formatting for the specified shape.
- `TextEffect As TextEffectFormat  (read-only)`  
  Returns a TextEffectFormat object that contains text-effect formatting properties for the specified shape. Read-only.
- `TextFrame As TextFrame  (read-only)`  
  Returns a TextFrame object that contains the text for the specified shape.
- `ThreeD As ThreeDFormat  (read-only)`  
  Returns a ThreeDFormat object that contains 3D formatting properties for the specified shape. Read-only.
- `Top As Single  (read/write)`  
  Returns or sets the vertical position of the specified shape or shape range in points. Read/write Single.
- `Type As MsoShapeType  (read-only)`  
  Returns the type of inline shape. Read-only MsoShapeType.
- `VerticalFlip As MsoTriState  (read-only)`  
  True if the specified shape is flipped around the vertical axis. Read-only MsoTriState.
- `Vertices As Variant  (read-only)`  
  Returns the coordinates of the specified freeform drawing's vertices (and control points for Bzier curves) as a series of coordinate pairs. Read-only Variant.
- `Visible As MsoTriState  (read/write)`  
  True if the specified object, or the formatting applied to it, is visible. Read/write MsoTriState.
- `Width As Single  (read/write)`  
  Returns or sets the width, in points, of the specified shape. Read/write Long.
- `ZOrderPosition As Long  (read-only)`  
  Returns a Long that represents the position of the specified shape in the z-order. Read-only.
- `Hyperlink As Hyperlink  (read-only)`  
  Returns a Hyperlink object that represents the hyperlink associated with a Shape object. Read-only.
- `RelativeHorizontalPosition As WdRelativeHorizontalPosition  (read/write)`  
  Specifies to the relative horizontal position of a shape. Read/write WdRelativeHorizontalPosition.
- `RelativeVerticalPosition As WdRelativeVerticalPosition  (read/write)`  
  Specifies the relative vertical position of a shape. Read/write WdRelativeVerticalPosition.
- `LockAnchor As Long  (read/write)`  
  True if the anchor of a Shape object is locked to the anchoring range. Read/write Long.
- `WrapFormat As WrapFormat  (read-only)`  
  Returns a WrapFormat object that contains the properties for wrapping text around the specified shape. Read-only.
- `OLEFormat As OLEFormat  (read-only)`  
  Returns an OLEFormat object that represents the OLE characteristics (other than linking) for the specified shape, inline shape, or field. Read-only.
- `Anchor As Range  (read-only)`  
  Returns a Range object that represents the anchoring range for the specified shape or shape range. Read-only.
- `LinkFormat As LinkFormat  (read-only)`  
  Returns a LinkFormat object that represents the link options of a shape that is linked to a file. Read/only.
- `AlternativeText As String  (read/write)`  
  Returns or sets the alternative text associated with a shape in a webpage. Read/write String.
- `Script As Script  (read-only)`  
  Returns a Script object, which represents a block of script or code for an image on a webpage.
- `Child As MsoTriState  (read-only)`  
  True if the shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only MsoTriState.
- `ParentGroup As Shape  (read-only)`  
  Returns a Shape object that represents the common parent shape of a child shape or a range of child shapes.
- `CanvasItems As CanvasShapes  (read-only)`  
  Returns a CanvasShapes object that represents a collection of shapes in a drawing canvas.
- `ID As Long  (read-only)`  
  Returns the identification type for the specified shape. Read-only Long.
- `LayoutInCell As Long  (read/write)`  
  Returns a Long that represents whether a shape in a table is displayed inside or outside the table.
- `HasChart As MsoTriState  (read-only)`  
  True if the specified shape has a chart. Read-only.
- `Chart As Chart  (read-only)`  
  Returns a Chart object that represents a chart within the collection of shapes in a document. Read-only.
- `LeftRelative As Single  (read/write)`  
  Returns or sets a Single that represents the relative left position of a shape. Read/write.
- `TopRelative As Single  (read/write)`  
  Returns or sets a Single that represents the relative top position of a shape. Read/write.
- `WidthRelative As Single  (read/write)`  
  Returns or sets a Single that represents the relative width of a shape. Read/write.
- `HeightRelative As Single  (read/write)`  
  Returns or sets a Single that represents the percentage of the relative height of a shape. Read/write.
- `RelativeHorizontalSize As WdRelativeHorizontalSize  (read/write)`  
  Returns or sets a WdRelativeVerticalSize constant that represents the object to which a range of shapes is relative. Read/write.
- `RelativeVerticalSize As WdRelativeVerticalSize  (read/write)`  
  Returns or sets a WdRelativeVerticalSize constant that represents the relative vertical size of a shape. Read/write.
- `SoftEdge As SoftEdgeFormat  (read-only)`  
  Returns a SoftEdgeFormat object that represents the soft edge formatting for a shape. Read-only.
- `Glow As GlowFormat  (read-only)`  
  Returns a GlowFormat object that represents the glow formatting for a shape. Read-only.
- `Reflection As ReflectionFormat  (read-only)`  
  Returns a ReflectionFormat object that represents the reflection formatting for a shape. Read-only.
- `TextFrame2 As TextFrame2  (read-only)`  
  Returns a TextFrame2 object that contains the text for the specified shape. Read-only.
- `HasSmartArt As MsoTriState  (read-only)`  
  Returns True if there is a SmartArt diagram present on the shape. Read-only.
- `SmartArt As SmartArt  (read-only)`  
  Returns a SmartArt object that provides a way to work with the SmartArt associated with the specified shape. Read-only.
- `ShapeStyle As MsoShapeStyleIndex  (read/write)`  
  Returns or sets the shape style for the specified shape. Read/write MsoShapeStyleIndex.
- `BackgroundStyle As MsoBackgroundStyleIndex  (read/write)`  
  Sets or returns the background style of the specified shape. Read/write MsoBackgroundStyleIndex.
- `Title As String  (read/write)`  
  Returns or sets a String that contains a title for the specified shape. Read/write.
- `GraphicStyle As MsoGraphicStyleIndex  (read/write)`  
  Returns or sets an MsoGraphicStyleIndex constant that represents the style of an SVG graphic. Read/write.
- `Model3D As Model3DFormat  (read-only)`  
  Returns a Model3DFormat object that contains Model3D properties. Read-only.
- `Decorative As MsoTriState  (read/write)`  
  Sets or returns the decorative flag for the specified object. Read/write.

## Methods (19)

- `Apply()`  
  Applies to the specified shape formatting that has been copied using the PickUp method.
- `Delete()`  
  Deletes the specified shape node.
- `Duplicate() As Shape`  
  Creates a duplicate of the specified Shape object, adds the new shape to the Shapes collection at a standard offset from the original shapes, and then returns the new Shape object.
- `Flip(FlipCmd As MsoFlipCmd)`  
  Flips a shape horizontally or vertically.
    - `FlipCmd As MsoFlipCmd` (required): The flip orientation.
- `IncrementLeft(Increment As Single)`  
  Moves the specified shape horizontally by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shape is to be moved horizontally, in points. A positive value moves the shape to the right; a negative value moves it to the left.
- `IncrementRotation(Increment As Single)`  
  Changes the rotation of the specified shape around the z-axis by the specified number of degrees.
    - `Increment As Single` (required): Specifies how far the shape is to be rotated horizontally, in degrees. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise.
- `IncrementTop(Increment As Single)`  
  Moves the specified shape vertically by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shape object is to be moved vertically, in points. A positive value moves the shape down; a negative value moves it up.
- `PickUp()`  
  Copies the formatting of the specified shape.
- `ScaleHeight(Factor As Single, RelativeToOriginalSize As MsoTriState, [Scale As MsoScaleFrom])`  
  Scales the height of the shape by a specified factor.
    - `Factor As Single` (required): Specifies the ratio between the height of the shape after you resize it and the current or original height. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.
    - `RelativeToOriginalSize As MsoTriState` (required): True to scale the shape relative to its original size. False to scale it relative to its current size. You can specify True for this argument only if the specified shape is a picture or an OLE object.
    - `Scale As MsoScaleFrom` (optional): The part of the shape that retains its position when the shape is scaled.
- `ScaleWidth(Factor As Single, RelativeToOriginalSize As MsoTriState, [Scale As MsoScaleFrom])`  
  Scales the width of the shape by a specified factor.
    - `Factor As Single` (required): Specifies the ratio between the width of the shape after you resize it and the current or original width. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.
    - `RelativeToOriginalSize As MsoTriState` (required): True to scale the shape relative to its original size. False to scale it relative to its current size. You can specify True for this argument only if the specified shape is a picture or an OLE object.
    - `Scale As MsoScaleFrom` (optional): The part of the shape that retains its position when the shape is scaled.
- `Select([Replace As Variant])`  
  Selects the specified shape.
    - `Replace As Variant` (optional): If adding a shape, True replaces the selection. False adds the new shape to the selection.
- `SetShapesDefaultProperties()`  
  Applies the formatting of the default shape for a document to the specified shape.
- `Ungroup() As ShapeRange`  
  Ungroups any grouped shapes in the specified shape.
- `ZOrder(ZOrderCmd As MsoZOrderCmd)`  
  Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).
    - `ZOrderCmd As MsoZOrderCmd` (required): Specifies where to move the specified shape relative to the other shapes.
- `ConvertToInlineShape() As InlineShape`  
  Converts the specified shape in the drawing layer of a document to an inline shape in the text layer. You can convert only shapes that represent pictures, OLE objects, or ActiveX controls. This method returns an InlineShape object that represents the picture or OLE object.
- `CanvasCropLeft(Increment As Single)`  
  Crops a percentage of the width of a drawing canvas from the left side of the canvas.
    - `Increment As Single` (required): The amount in percentage points of the drawing canvas's width that you want remaining after the canvas is cropped. Entering 0.9 as the increment crops ten percent of the canvas's width from the left. Entering 0.1 crops ninety percent of the canvas's width from the left.
- `CanvasCropTop(Increment As Single)`  
  Crops a percentage of the height of a drawing canvas from the top of the canvas.
    - `Increment As Single` (required): The amount in percentage points of a canvas's height that you want remaining after the canvas is cropped. Entering 0.9 as the increment crops ten percent of the canvas's height from the top. Entering 0.1 crops ninety percent of the canvas's height from the top.
- `CanvasCropRight(Increment As Single)`  
  Crops a percentage of the width of a drawing canvas from the right side of the canvas.
    - `Increment As Single` (required): The amount in percentage points of the canvas's width that you want remaining after the canvas is cropped. Entering 0.9 as the increment crops ten percent of the canvas's width from the right. Entering 0.1 crops ninety percent of the canvas's width from the right.
- `CanvasCropBottom(Increment As Single)`  
  Crops a percentage of the height of a drawing canvas from the bottom of the canvas.
    - `Increment As Single` (required): The amount in percentage points of a drawing canvas's height that you want remaining after the canvas is cropped. Entering 0.9 as the increment crops ten percent of the canvas's height from the bottom. Entering 0.1 crops ninety percent of the canvas's height from the bottom.
