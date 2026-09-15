# ShapeRange

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209B5-0000-0000-C000-000000000046}  

Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as one shape or as many as all the shapes in the document.

**Remarks:** You can include whichever shapes you want-chosen from among all the shapes in the document or all the shapes in the selection-to construct a shape range. For example, you could construct a ShapeRange collection that contains the first three shapes in a document, all the selected shapes in a document, or all the freeform shapes in a document. Most operations that you can do with a Shape object, you can also do with a ShapeRange object that contains only one shape. Some operations, when performed on a ShapeRange object that contains more than one shape, will cause an error. Use Range (index), where index is the name or index number of the shape or an array that contains either names or index numbers of shapes, to return a ShapeRange collection that represents a set of shapes on a document. Use Visual Basic's Array function to construct an array of names or index numbers. The following example sets the fill pattern for shapes one and three on the active document. The following example selects the shapes named Oval 4 and Rectangle 5 on the active document.

## Properties (58)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ShapeRange object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of shapes in the collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Adjustments As Adjustments  (read-only)`  
  Returns an Adjustments object that contains adjustment values for all the adjustments in the specified ShapeRange object that represents an AutoShape or WordArt. Read-only.
- `AutoShapeType As MsoAutoShapeType  (read/write)`  
  Returns or sets the shape type for the specified ShapeRange object, which must represent an AutoShape other than a line or freeform drawing. Read/write MsoAutoShapeType.
- `Callout As CalloutFormat  (read-only)`  
  Returns a CalloutFormat object that contains callout formatting properties for the specified shape. Read-only.
- `Fill As FillFormat  (read-only)`  
  Returns a FillFormat object that contains fill formatting properties for the specified shape. Read-only.
- `GroupItems As GroupShapes  (read-only)`  
  Returns a GroupShapes object that represents the individual shapes in the specified group. Read-only.
- `Height As Single  (read/write)`  
  Returns or sets the height of the specified shape range. Read/write Single.
- `HorizontalFlip As MsoTriState  (read-only)`  
  Indicates that a range of shapes has been flipped horizontally. Read-only MsoTriState.
- `Left As Single  (read/write)`  
  Returns or sets a Single that represents the horizontal position, measured in points, of the specified range of shapes. Can also be any valid WdShapePosition constant. Read/write.
- `Line As LineFormat  (read-only)`  
  Returns a LineFormat object that contains line formatting properties for the specified range of shapes. Read-only.
- `LockAspectRatio As MsoTriState  (read/write)`  
  MsoTrue if the specified shape retains its original proportions when you resize it. MsoFalse if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write String.
- `Nodes As ShapeNodes  (read-only)`  
  Returns a ShapeNodes collection that represents the geometric description of the specified shape.
- `Rotation As Single  (read/write)`  
  Returns or sets the number of degrees the specified shape is rotated around the z-axis. Read/write Single.
- `PictureFormat As PictureFormat  (read-only)`  
  Returns a PictureFormat object that contains picture formatting properties for the specified range of shapes. Read-only.
- `Shadow As ShadowFormat  (read-only)`  
  Returns a ShadowFormat object that represents the shadow formatting for the specified shape.
- `TextEffect As TextEffectFormat  (read-only)`  
  Returns a TextEffectFormat object that contains text-effect formatting properties for the specified shape. Read-only.
- `TextFrame As TextFrame  (read-only)`  
  Returns a TextFrame object that contains the text for the specified range of shapes.
- `ThreeD As ThreeDFormat  (read-only)`  
  Returns a ThreeDFormat object that contains 3D formatting properties for the specified range of shapes. Read-only.
- `Top As Single  (read/write)`  
  Returns or sets the vertical position of the specified shape or shape range in points. Read/write Single.
- `Type As MsoShapeType  (read-only)`  
  Returns the shape type. Read-only MsoShapeType.
- `VerticalFlip As MsoTriState  (read-only)`  
  True if the specified shape is flipped around the vertical axis. Read-only MsoTriState.
- `Vertices As Variant  (read-only)`  
  Returns the coordinates of the specified freeform drawing's vertices (and control points for Bzier curves) as a series of coordinate pairs. Use the array returned by this property as an argument for the AddCurve or AddPolyLine method. Read-only Variant.
- `Visible As MsoTriState  (read/write)`  
  True if the specified object, or the formatting applied to it, is visible. Read/write MsoTriState.
- `Width As Single  (read/write)`  
  Returns or sets the width, in points, of the shapes within the range. Read/write Long.
- `ZOrderPosition As Long  (read-only)`  
  Returns a Long that represents the position of the specified shape in the z-order. Read-only.
- `Hyperlink As Hyperlink  (read-only)`  
  Returns a Hyperlink object that represents the hyperlink associated with the specified ShapeRange object. Read-only.
- `RelativeHorizontalPosition As WdRelativeHorizontalPosition  (read/write)`  
  Specifies the relative horizontal position of a range of shapes. Read/write WdRelativeHorizontalPosition.
- `RelativeVerticalPosition As WdRelativeVerticalPosition  (read/write)`  
  Specifies the relative vertical position of a range of shapes. Read/write WdRelativeHorizontalPosition.
- `LockAnchor As Long  (read/write)`  
  True if the anchor for the specified ShapeRange object is locked to the anchoring range. Read/write Long.
- `WrapFormat As WrapFormat  (read-only)`  
  Returns a WrapFormat object that contains the properties for wrapping text around the specified range of shapes. Read-only.
- `Anchor As Range  (read-only)`  
  Returns a Range object that represents the anchoring range for the specified shape range. Read-only.
- `AlternativeText As String  (read/write)`  
  Returns or sets the alternative text associated with a shape in a webpage. Read/write String.
- `Child As MsoTriState  (read-only)`  
  True if the shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only MsoTriState.
- `ParentGroup As Shape  (read-only)`  
  Returns a Shape object that represents the common parent shape of a range of shapes.
- `CanvasItems As CanvasShapes  (read-only)`  
  Returns a CanvasShapes object that represents a collection of shapes in a drawing canvas.
- `ID As Long  (read-only)`  
  Returns the identification type for the range of shapes. Read-only Long.
- `LayoutInCell As Long  (read/write)`  
  Returns a Long that represents whether a shape in a table is displayed inside the table or outside the table. .
- `LeftRelative As Single  (read/write)`  
  Returns or sets a Single that represents the relative left position of a range of shapes. Read/write.
- `TopRelative As Single  (read/write)`  
  Returns or sets a Single that represents the relative top position of a range of shapes. Read/write.
- `WidthRelative As Single  (read/write)`  
  Returns or sets a Single that represents the relative width of a range of shapes. Read/write.
- `HeightRelative As Single  (read/write)`  
  Returns or sets a Single that represents the percentage of the target shape to which the range of shapes is sized. Read/write.
- `RelativeHorizontalSize As WdRelativeHorizontalSize  (read/write)`  
  Returns or sets a WdRelativeHorizontalSize constant that represents the object to which a range of shapes is relative. Read/write.
- `RelativeVerticalSize As WdRelativeVerticalSize  (read/write)`  
  Returns or sets a WdRelativeVerticalSize constant that represents the object to which a range of shapes is relative. Read/write.
- `SoftEdge As SoftEdgeFormat  (read-only)`  
  Returns a SoftEdgeFormat object that represents the soft edge formatting for a range of shapes. Read-only.
- `Glow As GlowFormat  (read-only)`  
  Returns a GlowFormat object that represents the glow formatting for a range of shapes. Read-only.
- `Reflection As ReflectionFormat  (read-only)`  
  Returns a ReflectionFormat object that represents the reflection formatting for a range of shapes. Read-only.
- `TextFrame2 As TextFrame2  (read-only)`  
  Returns a TextFrame2 object that contains the text for the specified range of shapes. Read-only.
- `ShapeStyle As MsoShapeStyleIndex  (read/write)`  
  Returns or sets the shape style for the shapes in the specified shape range. Read/write MsoShapeStyleIndex.
- `BackgroundStyle As MsoBackgroundStyleIndex  (read/write)`  
  Sets or returns the background style of the shapes in the specified shape range. Read/write MsoBackgroundStyleIndex.
- `Title As String  (read/write)`  
  Returns or sets a String that contains a title for the shapes in the specified shape range. Read/write.
- `GraphicStyle As MsoGraphicStyleIndex  (read/write)`  
  Returns or sets an MsoGraphicStyleIndex constant that represents the style of a shape range containing one or more SVG graphics. Read/write.
- `Model3D As Model3DFormat  (read-only)`  
  Returns a Model3DFormat object that contains Model3D properties. Read-only.
- `Decorative As MsoTriState  (read/write)`  
  Sets or returns the decorative flag for the specified object. Read/write.

## Methods (23)

- `Item(Index As Variant) As Shape`  
  Returns an individual Shape object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Align(Align As MsoAlignCmd, RelativeTo As Long)`  
  Aligns the shapes in the specified range of shapes.
    - `Align As MsoAlignCmd` (required): Specifies the way the shapes in the specified shape range are to be aligned.
    - `RelativeTo As Long` (required): True to align shapes relative to the edge of the document. False to align shapes relative to one another.
- `Apply()`  
  Applies to the specified shape formatting that has been copied using the PickUp method.
- `Delete()`  
  Deletes the specified range of shapes.
- `Distribute(Distribute As MsoDistributeCmd, RelativeTo As Long)`  
  Evenly distributes the shapes in the specified range of shapes. .
    - `Distribute As MsoDistributeCmd` (required): Specifies whether to distribute shapes horizontally or vertically.
    - `RelativeTo As Long` (required): True to distribute the shapes evenly over the entire horizontal or vertical space on the page. False to distribute them within the horizontal or vertical space that the range of shapes originally occupies.
- `Duplicate() As ShapeRange`  
  Creates a duplicate of the specified ShapeRange object, adds the new range of shapes to the Shapes collection at a standard offset from the original shapes, and then returns a Shape object.
- `Flip(FlipCmd As MsoFlipCmd)`  
  Flips a shape horizontally or vertically.
    - `FlipCmd As MsoFlipCmd` (required): The flip orientation.
- `IncrementLeft(Increment As Single)`  
  Moves the specified shape horizontally by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shape is to be moved horizontally, in points. A positive value moves the shape to the right; a negative value moves it to the left.
- `IncrementRotation(Increment As Single)`  
  Changes the rotation of the specified shape around the z-axis by the specified number of degrees. .
    - `Increment As Single` (required): Specifies how far the shape is to be rotated horizontally, in degrees. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise.
- `IncrementTop(Increment As Single)`  
  Moves the specified shape vertically by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shape object is to be moved vertically, in points. A positive value moves the shape down; a negative value moves it up.
- `Group() As Shape`  
  Groups the shapes in the specified range, and returns the grouped shapes as a single Shape object.
- `PickUp()`  
  Copies the formatting of the specified shape.
- `ScaleHeight(Factor As Single, RelativeToOriginalSize As MsoTriState, [Scale As MsoScaleFrom])`  
  Scales the height of a range of shapes by a specified factor.
    - `Factor As Single` (required): Specifies the ratio between the height of the shape after you resize it and the current or original height. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.
    - `RelativeToOriginalSize As MsoTriState` (required): True to scale the shape relative to its original size. False to scale it relative to its current size. You can specify True for this argument only if the specified shape is a picture or an OLE object.
    - `Scale As MsoScaleFrom` (optional): The part of the shape that retains its position when the shape is scaled.
- `ScaleWidth(Factor As Single, RelativeToOriginalSize As MsoTriState, [Scale As MsoScaleFrom])`  
  Scales the width of a shape by a specified factor.
    - `Factor As Single` (required): Specifies the ratio between the width of the shape after you resize it and the current or original width. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.
    - `RelativeToOriginalSize As MsoTriState` (required): True to scale the shape relative to its original size. False to scale it relative to its current size. You can specify True for this argument only if the specified shape is a picture or an OLE object.
    - `Scale As MsoScaleFrom` (optional): The part of the shape that retains its position when the shape is scaled.
- `Select([Replace As Variant])`  
  Selects the specified range of shapes.
    - `Replace As Variant` (optional): If adding a shape, True replaces the selection. False adds the new shape to the selection.
- `SetShapesDefaultProperties()`  
  Applies the formatting of a default shape for a document to the specified range of shapes.
- `Ungroup() As ShapeRange`  
  Ungroups any grouped shapes in the specified range of shapes, disassembles pictures and OLE objects within the specified shape or range of shapes, and returns the ungrouped shapes as a single ShapeRange object.
- `ZOrder(ZOrderCmd As MsoZOrderCmd)`  
  Moves the specified shape range in front of or behind other shapes in the collection (that is, changes the shape range's position in the z-order).
    - `ZOrderCmd As MsoZOrderCmd` (required): Specifies where to move the specified shape range relative to the other shapes.
- `ConvertToInlineShape() As InlineShape`  
  Converts the specified shape in the drawing layer of a document to an inline shape in the text layer. You can convert only shapes that represent pictures, OLE objects, or ActiveX controls. .
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
