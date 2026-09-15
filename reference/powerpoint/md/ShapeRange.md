# ShapeRange

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149347A-5A91-11CF-8700-00AA0060263B}  

Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as a single shape or as many as all the shapes on the document.

**Remarks:** You can include whichever shapes you want-chosen from among all the shapes on the document or all the shapes in the selection-to construct a shape range. For example, you could construct a ShapeRange collection that contains the first three shapes on a document, all the selected shapes on a document, or all the freeforms on a document. For an overview of how to work with either a single shape or with more than one shape at a time, see Work with shapes (drawing objects). The following examples describe how to: - Return a set of shapes that you specify by name or index number. - Return all or some of the selected shapes on a document.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.Range(Array(1, 3)).Fill _

    .Patterned msoPatternHorizontalBrick
```

## Properties (68)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Adjustments As Adjustments  (read-only)`  
  Returns an Adjustments object that contains adjustment values for all the adjustments in the specified shape. Applies to any ShapeRange object that represents an AutoShape, WordArt, or a connector. Read-only.
- `AutoShapeType As MsoAutoShapeType  (read/write)`  
  Returns or sets the shape type for the specified ShapeRange object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write.
- `BlackWhiteMode As MsoBlackWhiteMode  (read/write)`  
  Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write.
- `Callout As CalloutFormat  (read-only)`  
  Returns a CalloutFormat object that contains callout formatting properties for the specified shape. Applies to Shape or ShapeRange objects that represent line callouts. Read-only.
- `ConnectionSiteCount As Long  (read-only)`  
  Returns the number of connection sites on the specified shape. Read-only.
- `Connector As MsoTriState  (read-only)`  
  Determines whether the specified shape is a connector. Read-only.
- `ConnectorFormat As ConnectorFormat  (read-only)`  
  Returns a ConnectorFormat object that contains connector formatting properties. Applies to Shape or ShapeRange objects that represent connectors. Read-only.
- `Fill As FillFormat  (read-only)`  
  Returns a FillFormat object that contains fill formatting properties for the specified shape. Read-only.
- `GroupItems As GroupShapes  (read-only)`  
  Returns a GroupShapes object that represents the individual shapes in the specified group. Use the Item method of the GroupShapes object to return a single shape from the group. Read-only.
- `Height As Single  (read/write)`  
  Returns or sets the height of the specified object, in points. Read/write.
- `HorizontalFlip As MsoTriState  (read-only)`  
  Returns whether the specified shape is flipped around the horizontal axis. Read-only.
- `Left As Single  (read/write)`  
  Returns or sets a Single that represents the distance in points from the left edge of the leftmost shape in the shape range to the left edge of the slide. Read/write.
- `Line As LineFormat  (read-only)`  
  Returns a LineFormat object that contains line formatting properties for the specified shape. (For a line, the LineFormat object represents the line itself; for a shape with a border, the LineFormat object represents the border.) Read-only.
- `LockAspectRatio As MsoTriState  (read/write)`  
  Determines whether the specified shape retains its original proportions when you resize it. Read/write.
- `Name As String  (read/write)`  
  When a shape is created, Microsoft PowerPoint automatically assigns it a name in the form _ShapeType Number_, where _ShapeType_ identifies the type of shape or AutoShape, and _Number_ is an integer that's unique within the collection of shapes on the slide. For example, the automatically generated names of the shapes on a slide could be Placeholder 1, Oval 2, and Rectangle 3. To avoid conflict with automatically assigned names, don't use the form _ShapeType Number_ for user-defined names, where _ShapeType_ is a value that is used for automatically generated names, and _Number_ is any positive integer. A shape range must contain exactly one shape. Read/write.
- `Nodes As ShapeNodes  (read-only)`  
  Returns a ShapeNodes collection that represents the geometric description of the specified shape. Applies to ShapeRange objects that represent freeform drawings.
- `Rotation As Single  (read/write)`  
  Returns or sets the number of degrees the specified shape is rotated around the z-axis. Read/write.
- `PictureFormat As PictureFormat  (read-only)`  
  Returns a PictureFormat object that contains picture formatting properties for the specified shape. Read-only.
- `Shadow As ShadowFormat  (read-only)`  
  Returns a ShadowFormat object that contains shadow formatting properties for the specified shapes. Read-only.
- `TextEffect As TextEffectFormat  (read-only)`  
  Returns a TextEffectFormat object that contains text-effect formatting properties for the specified shape. Read-only.
- `TextFrame As TextFrame  (read-only)`  
  Returns a TextFrame object that contains the alignment and anchoring properties for the specified shape or master text style. Read-only.
- `ThreeD As ThreeDFormat  (read-only)`  
  Returns a ThreeDFormat object that contains 3D - effect formatting properties for the specified shape. Read-only.
- `Top As Single  (read/write)`  
  Returns or sets a Single that represents the distance from the top edge of the topmost shape in the shape range to the top edge of the document. Read/write.
- `Type As MsoShapeType  (read-only)`  
  Represents the type of shape or shapes in a range of shapes. Read-only.
- `VerticalFlip As MsoTriState  (read-only)`  
  Determines whether the specified shape is flipped around the vertical axis. Read-only.
- `Vertices As Variant  (read-only)`  
  Returns the coordinates of the specified freeform drawing's vertices (and control points for Bzier curves) as a series of coordinate pairs. Read-only.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.
- `Width As Single  (read/write)`  
  Returns or sets the width of the specified object, in points. Read/write.
- `ZOrderPosition As Long  (read-only)`  
  Returns the position of the specified shape in the z-order. Read-only.
- `OLEFormat As OLEFormat  (read-only)`  
  Returns an OLEFormat object that contains OLE formatting properties for the specified shape. Applies to Shape or ShapeRange objects that represent OLE objects. Read-only.
- `LinkFormat As LinkFormat  (read-only)`  
  Returns a LinkFormat object that contains the properties that are unique to linked OLE objects. Read-only.
- `PlaceholderFormat As PlaceholderFormat  (read-only)`  
  Returns a PlaceholderFormat object that contains the properties that are unique to placeholders. Read-only.
- `AnimationSettings As AnimationSettings  (read-only)`  
  Returns an AnimationSettings object that represents all the special effects you can apply to the animation of the specified shape. Read-only.
- `ActionSettings As ActionSettings  (read-only)`  
  Returns an ActionSettings object that contains information about what action occurs when the user clicks or moves the mouse over the specified shape or text range during a slide show. Read-only.
- `Tags As Tags  (read-only)`  
  Returns a Tags object that represents the tags for the specified object. Read-only.
- `MediaType As PpMediaType  (read-only)`  
  Returns the OLE media type. Read-only.
- `HasTextFrame As MsoTriState  (read-only)`  
  Returns whether the specified shape has a text frame. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `AlternativeText As String  (read/write)`  
  Returns or sets the alternative text associated with a shape in a Web presentation. Read/write.
- `HasTable As MsoTriState  (read-only)`  
  Returns whether the specified shape is a table. Read-only.
- `Table As Table  (read-only)`  
  Returns a Table object that represents a table in a shape or in a shape range. Read-only.
- `Child As MsoTriState  (read-only)`  
  MsoTrue if the shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only.
- `ParentGroup As Shape  (read-only)`  
  Returns a Shape object that represents the common parent shape of a child shape or a range of child shapes.
- `Id As Long  (read-only)`  
  Returns a Long that identifies the shape or range of shapes. Read-only.
- `CustomerData As CustomerData  (read-only)`  
  Returns a CustomerData object.
- `TextFrame2 As TextFrame2  (read-only)`  
  Returns the TextFrame2 object associated with the specified ShapeRange object that contains the alignment and anchoring properties for the specified shape range. Read-only.
- `HasChart As MsoTriState  (read-only)`  
  Returns whether the shape range represented by the specified object contains a chart. Read-only.
- `ShapeStyle As MsoShapeStyleIndex  (read/write)`  
  Sets or returns the shape style index for the specified object.
- `BackgroundStyle As MsoBackgroundStyleIndex  (read/write)`  
  Sets or returns the background style of the specified object. Read/write.
- `SoftEdge As SoftEdgeFormat  (read-only)`  
  Returns the soft edge format for the specified range of shapes. Read-only.
- `Glow As GlowFormat  (read-only)`  
  Returns the glow format for the specified range of shapes. Read-only.
- `Reflection As ReflectionFormat  (read-only)`  
  Returns the reflection format for the specified range of shapes. Read-only.
- `Chart As Chart  (read-only)`  
  Returns the Chart object of the current ShapeRange object. Read-only.
- `HasSmartArt As MsoTriState  (read-only)`  
  Returns True if the current ShapeRange object has a SmartArt diagram. Read-only.
- `SmartArt As SmartArt  (read-only)`  
  Returns the SmartArt diagram of the ShapeRange object. Read-only.
- `Title As String  (read/write)`  
  Returns a Shape object that represents the slide title. Read-only.
- `MediaFormat As MediaFormat  (read-only)`  
  Returns the current MediaFormat object. Read-only.
- `IsNarration As MsoTriState  (read/write)`  
  Specifies whether the specified shape range contains a narration. Read/write.
- `InkXML As String  (read-only)`  
  Returns a String that contains the InkActionML associated with the specified shape range. Read-only.
- `HasInkXML As MsoTriState  (read-only)`  
  Returns an MsoTriState enumeration value that indicates whether the specified shape range contains ink XML that can be retrieved via the ShapeRange.InkXML property. Read-only.
- `HasSectionZoom As MsoTriState  (read-only)`
- `GraphicStyle As MsoGraphicStyleIndex  (read/write)`  
  Returns or sets an MsoGraphicStyleIndex constant that represents the style of a shape range containing one or more SVG graphics. Read/write.
- `Model3D As Model3DFormat  (read-only)`  
  Returns a Model3DFormat object that represents the 3D properties of a 3D model object. Read-only.
- `Decorative As MsoTriState  (read/write)`  
  Sets or returns the decorative flag for the specified object. Read/write.
- `Locked As MsoTriState  (read/write)`

## Methods (28)

- `Apply()`  
  Applies to the specified shape range formatting that's been copied by using the PickUp method.
- `Delete()`  
  Deletes the specified ShapeRange object.
- `Flip(FlipCmd As MsoFlipCmd)`  
  Flips the specified shape range around its horizontal or vertical axis.
    - `FlipCmd As MsoFlipCmd` (required): Specifies whether the shape is to be flipped horizontally or vertically.
- `IncrementLeft(Increment As Single)`  
  Moves the specified shape range horizontally by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shape range is to be moved horizontally, in points. A positive value moves the shape range to the right; a negative value moves it to the left.
- `IncrementRotation(Increment As Single)`  
  Changes the rotation of the specified shape range around the z-axis by the specified number of degrees. Use the Rotation property to set the absolute rotation of the shape range.
    - `Increment As Single` (required): Specifies how far the shape range is to be rotated horizontally, in degrees. A positive value rotates the shape range clockwise; a negative value rotates it counterclockwise.
- `IncrementTop(Increment As Single)`  
  Moves the specified shape range vertically by the specified number of points.
    - `Increment As Single` (required): Specifies how far the shape range is to be moved vertically, in points. A positive value moves the shape range down; a negative value moves it up.
- `PickUp()`  
  Copies the formatting of the specified shape. Use the Apply method to apply the copied formatting to another shape.
- `RerouteConnections()`  
  Reroutes connectors so that they take the shortest possible path between the shapes they connect. To do this, the RerouteConnections method may detach the ends of a connector and reattach them to different connecting sites on the connected shapes.
- `ScaleHeight(Factor As Single, RelativeToOriginalSize As MsoTriState, [fScale As MsoScaleFrom])`  
  Scales the height of the shapes in the range by a specified factor.
    - `Factor As Single` (required): Specifies the ratio between the height of the shapes after you resize them and their current or original height. For example, to make shapes 50 percent larger, specify 1.5 for this parameter.
    - `RelativeToOriginalSize As MsoTriState` (required): Specifies whether shapes are scaled relative to their current or original sizes.
    - `fScale As MsoScaleFrom` (optional): The parts of the shapes that retain their position when the shapes are scaled.
- `ScaleWidth(Factor As Single, RelativeToOriginalSize As MsoTriState, [fScale As MsoScaleFrom])`  
  Scales the width of the shapes in the range by a specified factor.
    - `Factor As Single` (required): Specifies the ratio between the width of the shapes after you resize them and the current or original width. For example, to make all the shapes in the range 50 percent larger, specify 1.5 for this parameter.
    - `RelativeToOriginalSize As MsoTriState` (required): Specifies whether shapes are scaled relative to their current or original size.
    - `fScale As MsoScaleFrom` (optional): The parts of the shapes that retain their positions when the shapes are scaled.
- `SetShapesDefaultProperties()`  
  Applies the formatting for the specified shape to the default shape. Shapes created after this method has been used will have this formatting applied to them by default.
- `Ungroup() As ShapeRange`  
  Ungroups any grouped shapes in the specified shape or range of shapes. Disassembles pictures and OLE objects within the specified shape or range of shapes. Returns the ungrouped shapes as a single ShapeRange object.
- `ZOrder(ZOrderCmd As MsoZOrderCmd)`  
  Moves the specified shape range in front of or behind other shapes in the collection (that is, changes the shape range's position in the z-order).
    - `ZOrderCmd As MsoZOrderCmd` (required): Specifies where to move the specified shape range relative to the other shapes.
- `Cut()`  
  Deletes the specified object and places it on the Clipboard.
- `Copy()`  
  Copies the specified object to the Clipboard.
- `Select([Replace As MsoTriState])`  
  Selects the specified object.
    - `Replace As MsoTriState` (optional): Specifies whether the selection replaces any previous selection.
- `Duplicate() As ShapeRange`  
  Creates a duplicate of the specified ShapeRange object, adds the range of shapes to the Shapes collection, and then returns the new ShapeRange object. The duplicated objects are placed at the end of the Shapes collection.
- `Item(Index As Variant) As Shape`  
  Returns a single Shape object from the specified ShapeRange collection.
    - `Index As Variant` (required): The name or index number of the single Shape object in the collection to be returned.
- `Group() As Shape`  
  Groups the shapes in the specified range. Returns the grouped shapes as a single Shape object.
- `Regroup() As Shape`  
  Regroups the group that the specified shape range belonged to previously. Returns the regrouped shapes as a single Shape object.
- `Align(AlignCmd As MsoAlignCmd, RelativeTo As MsoTriState)`  
  Aligns the shapes in the specified range of shapes.
    - `AlignCmd As MsoAlignCmd` (required): Specifies the way the shapes in the specified shape range are to be aligned.
    - `RelativeTo As MsoTriState` (required): Determines whether shapes are aligned relative to the edge of the slide.
- `Distribute(DistributeCmd As MsoDistributeCmd, RelativeTo As MsoTriState)`  
  Evenly distributes the shapes in the specified range of shapes. You can specify whether you want to distribute the shapes horizontally or vertically and whether you want to distribute them over the entire slide or just over the space they originally occupy.
    - `DistributeCmd As MsoDistributeCmd` (required): Specifies whether shapes in the range are to be distributed horizontally or vertically.
    - `RelativeTo As MsoTriState` (required): Determines whether shapes are distributed evenly over the entire horizontal or vertical space on the slide.
- `ConvertTextToSmartArt(Layout As SmartArtLayout)`  
  Returns the number of objects in the specified collection. Read-only.
- `PickupAnimation()`  
  Picks up all animation from the ShapeRange object.
- `ApplyAnimation()`  
  Applies animation to a ShapeRange object.
- `UpgradeMedia()`  
  Upgrades media within the specified ShapeRange object.
- `MergeShapes(MergeCmd As MsoMergeCmd, [PrimaryShape As Shape])`  
  Combines all of the shapes in the range to create a single new shape.
- `ConvertInkToText() As String`
