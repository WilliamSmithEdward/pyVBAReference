# ShapeRange

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002443B-0000-0000-C000-000000000046}  

Represents a shape range, which is a set of shapes on a document.

**Remarks:** A shape range can contain as few as a single shape or as many as all the shapes on the document. You can include whichever shapes you want-chosen from among all the shapes on the document or all the shapes in the selection-to construct a shape range. For example, you could construct a ShapeRange collection that contains the first three shapes on a document, all the selected shapes on a document, or all the freeforms on a document.

**Example:**

```vba
Set myDocument = Worksheets(1)
myDocument.Shapes.Range(Array(1, 3)).Fill.Patterned _
 msoPatternHorizontalBrick
```

## Properties (50)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `_NewEnum As IUnknown  (read-only)`
- `Adjustments As Adjustments  (read-only)`  
  Returns an Adjustments object that contains adjustment values for all the adjustments in the specified shape. Applies to any ShapeRange object that represents an AutoShape, WordArt, or Connector.
- `TextFrame As TextFrame  (read-only)`  
  Returns a TextFrame object that contains the alignment and anchoring properties for the specified shape. Read-only.
- `AutoShapeType As MsoAutoShapeType  (read/write)`  
  Returns or sets the shape type for the specified Shape or ShapeRange object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write MsoAutoShapeType.
- `Callout As CalloutFormat  (read-only)`  
  Returns a CalloutFormat object that contains callout formatting properties for the specified shape. Applies to ShapeRange objects that represent line callouts. Read-only.
- `ConnectionSiteCount As Long  (read-only)`  
  Returns the number of connection sites on the specified shape. Read-only Long.
- `Connector As MsoTriState  (read-only)`  
  True if the specified shape is a connector. Read-only MsoTriState.
- `ConnectorFormat As ConnectorFormat  (read-only)`  
  Returns a ConnectorFormat object that contains connector formatting properties. Applies to ShapeRange objects that represent connectors. Read-only.
- `Fill As FillFormat  (read-only)`  
  Returns a FillFormat object for a specified shape or a ChartFillFormat object for a specified chart that contains fill formatting properties for the shape or chart. Read-only.
- `GroupItems As GroupShapes  (read-only)`  
  Returns a GroupShapes object that represents the individual shapes in the specified group. Use the Item method of the GroupShapes object to return a single shape from the group. Applies to ShapeRange objects that represent grouped shapes. Read-only.
- `Height As Single  (read/write)`  
  Returns or sets a Single value that represents the height, in points, of the object.
- `HorizontalFlip As MsoTriState  (read-only)`  
  True if the specified shape is flipped around the horizontal axis. Read-only MsoTriState.
- `Left As Single  (read/write)`  
  Returns or sets a Single value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).
- `Line As LineFormat  (read-only)`  
  Returns a LineFormat object that contains line formatting properties for the specified shape. (For a line, the LineFormat object represents the line itself; for a shape with a border, the LineFormat object represents the border). Read-only.
- `LockAspectRatio As MsoTriState  (read/write)`  
  True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `Nodes As ShapeNodes  (read-only)`  
  Returns a ShapeNodes collection that represents the geometric description of the specified shape.
- `Rotation As Single  (read/write)`  
  Returns or sets the rotation of the shape, in degrees. Read/write Single.
- `PictureFormat As PictureFormat  (read-only)`  
  Returns a PictureFormat object that contains picture formatting properties for the specified shape. Applies to a ShapeRange object that represent pictures or OLE objects. Read-only.
- `Shadow As ShadowFormat  (read-only)`  
  Returns a read-only ShadowFormat object that contains shadow formatting properties for the specified shape or shapes.
- `TextEffect As TextEffectFormat  (read-only)`  
  Returns a TextEffectFormat object that contains text-effect formatting properties for the specified shape. Read-only.
- `ThreeD As ThreeDFormat  (read-only)`  
  Returns a ThreeDFormat object that contains 3D-effect formatting properties for the specified shape. Read-only.
- `Top As Single  (read/write)`  
  Returns or sets a Single value that represents the distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.
- `Type As MsoShapeType  (read-only)`  
  Returns an MsoShapeType value that represents the shape type.
- `VerticalFlip As MsoTriState  (read-only)`  
  True if the specified shape is flipped around the vertical axis. Read-only MsoTriState.
- `Vertices As Variant  (read-only)`  
  Returns the coordinates of the specified freeform drawing's vertices (and control points for Bzier curves) as a series of coordinate pairs. Use the array returned by this property as an argument to the AddCurve method or AddPolyLine method. Read-only Variant.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets an MsoTriState value that determines whether the object is visible. Read/write.
- `Width As Single  (read/write)`  
  Returns or sets a Single value that represents the width, in points, of the object.
- `ZOrderPosition As Long  (read-only)`  
  Returns the position of the specified shape in the z-order. Read-only Long.
- `BlackWhiteMode As MsoBlackWhiteMode  (read/write)`  
  Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write MsoBlackWhiteMode.
- `AlternativeText As String  (read/write)`  
  Returns or sets the descriptive (alternative) text string for a ShapeRange object when the object is saved to a webpage. Read/write String.
- `Child As MsoTriState  (read-only)`  
  Returns msoTrue if the specified shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Returns msoFalse if the selected shape is not a child shape. Returns msoTriStateMixed if only some of the selected shapes are child shapes. Read-only MsoTriState.
- `ParentGroup As Shape  (read-only)`  
  Returns a Shape object that represents the common parent shape of a child shape or a range of child shapes.
- `ID As Long  (read-only)`  
  Returns a Long value that represents the type for the specified object.
- `Chart As Chart  (read-only)`  
  Returns a Chart object that represents the chart contained in the shape range. Read-only.
- `HasChart As MsoTriState  (read-only)`  
  Returns whether a shape range contains a chart. Read-only MsoTriState.
- `TextFrame2 As TextFrame2  (read-only)`  
  Returns a TextFrame2 object that contains text formatting for the specified shape range. Read-only.
- `ShapeStyle As MsoShapeStyleIndex  (read/write)`  
  Returns or sets an MsoShapeStyleIndex value that represents the shape style of the shape range. Read/write.
- `BackgroundStyle As MsoBackgroundStyleIndex  (read/write)`  
  Returns or sets the background style. Read/write MsoBackgroundStyleIndex.
- `SoftEdge As SoftEdgeFormat  (read-only)`  
  Returns a SoftEdgeFormat object for a specified shape range that contains soft edge formatting properties for the shape range. Read-only.
- `Glow As GlowFormat  (read-only)`  
  Returns a GlowFormat object for a specified shape range that contains glow formatting properties for the shape range. Read-only.
- `Reflection As ReflectionFormat  (read-only)`  
  Returns a ReflectionFormat object for a specified shape range that contains reflection formatting properties for the shape range. Read-only.
- `Title As String  (read/write)`  
  Returns or sets the title of the alternative text associated with the specified shape range. Read/write.
- `GraphicStyle As MsoGraphicStyleIndex  (read/write)`  
  Returns or sets an MsoGraphicStyleIndex constant that represents the style of a shape range containing one or more SVG graphics. Read/write.
- `Model3D As Model3DFormat  (read-only)`  
  Returns a Model3DFormat object that contains Model3D properties. Read-only.
- `Decorative As MsoTriState  (read/write)`  
  Sets or returns the decorative flag for the specified object. Read/write.

## Methods (21)

- `Item(Index As Variant) As Shape`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_Default(Index As Variant) As Shape`
- `Align(AlignCmd As MsoAlignCmd, RelativeTo As MsoTriState)`  
  Aligns the shapes in the specified range of shapes.
    - `AlignCmd As MsoAlignCmd` (required): Specifies the way that the shapes in the specified shape range are to be aligned.
    - `RelativeTo As MsoTriState` (required): Not used in Microsoft Excel. Must be False.
- `Apply()`  
  Applies to the specified shape formatting that's been copied by using the PickUp method.
- `Delete()`  
  Deletes the object.
- `Distribute(DistributeCmd As MsoDistributeCmd, RelativeTo As MsoTriState)`  
  Horizontally or vertically distributes the shapes in the specified range of shapes.
    - `DistributeCmd As MsoDistributeCmd` (required): Specifies whether shapes in the range are to be distributed horizontally or vertically.
    - `RelativeTo As MsoTriState` (required): Not used in Microsoft Excel. Must be False.
- `Duplicate() As ShapeRange`  
  Duplicates the object and returns a reference to the new copy.
- `Flip(FlipCmd As MsoFlipCmd)`  
  Flips the specified shape around its horizontal or vertical axis.
    - `FlipCmd As MsoFlipCmd` (required): Specifies whether the shape is to be flipped horizontally or vertically.
- `IncrementLeft(Increment As Single)`  
  Moves the specified shape horizontally by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shape is to be moved horizontally, in points. A positive value moves the shape to the right; a negative value moves it to the left.
- `IncrementRotation(Increment As Single)`  
  Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the Rotation property to set the absolute rotation of the shape.
    - `Increment As Single` (required): Specifies how far the shape is to be rotated horizontally, in degrees. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise.
- `IncrementTop(Increment As Single)`  
  Moves the specified shape vertically by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shape object is to be moved vertically, in points. A positive value moves the shape down; a negative value moves it up.
- `Group() As Shape`  
  Groups the shapes in the specified range.
- `PickUp()`  
  Copies the formatting of the specified shape. Use the Apply method to apply the copied formatting to another shape.
- `RerouteConnections()`  
  This method reroutes all connectors attached to the specified shape; if the specified shape is a connector, it's rerouted.
- `Regroup() As Shape`  
  Regroups the group that the specified shape range belonged to previously. Returns the regrouped shapes as a single Shape object.
- `ScaleHeight(Factor As Single, RelativeToOriginalSize As MsoTriState, [Scale As Variant])`  
  Scales the height of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current height.
    - `Factor As Single` (required): Specifies the ratio between the height of the shape after you resize it and the current or original height. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.
    - `RelativeToOriginalSize As MsoTriState` (required): msoTrue to scale the shape relative to its original size. msoFalse to scale it relative to its current size. You can specify msoTrue for this argument only if the specified shape is a picture or an OLE object.
    - `Scale As Variant` (optional): One of the constants of MsoScaleFrom, which specifies which part of the shape retains its position when the shape is scaled.
- `ScaleWidth(Factor As Single, RelativeToOriginalSize As MsoTriState, [Scale As Variant])`  
  Scales the width of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.
    - `Factor As Single` (required): Specifies the ratio between the width of the shape after you resize it and the current or original width. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.
    - `RelativeToOriginalSize As MsoTriState` (required): msoFalse to scale it relative to its current size. You can specify msoTrue for this argument only if the specified shape is a picture or an OLE object.
    - `Scale As Variant` (optional): One of the constants of MsoScaleFrom, which specifies which part of the shape retains its position when the shape is scaled.
- `Select([Replace As Variant])`  
  Selects the object.
    - `Replace As Variant` (optional): (used only with sheets). True to replace the current selection with the specified object. False to extend the current selection to include any previously selected objects and the specified object.
- `SetShapesDefaultProperties()`  
  Makes the formatting of the specified shape the default formatting for the shape.
- `Ungroup() As ShapeRange`  
  Ungroups any grouped shapes in the specified shape or range of shapes. Disassembles pictures and OLE objects within the specified shape or range of shapes.
- `ZOrder(ZOrderCmd As MsoZOrderCmd)`  
  Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).
    - `ZOrderCmd As MsoZOrderCmd` (required): Specifies where to move the specified shape relative to the other shapes.
