# InlineShape

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209A8-0000-0000-C000-000000000046}  

Represents an object in the text layer of a document. An inline shape can only be a picture, an OLE object, or an ActiveX control. The InlineShape object is a member of the InlineShapes collection. The InlineShapes collection contains all the shapes that appear inline in a document, range, or selection.

**Remarks:** InlineShape objects are treated like characters and are positioned as characters within a line of text. Use InlineShapes (Index), where Index is the index number, to return a single InlineShape object. Inline shapes don't have names. The following example activates the first inline shape in the active document. Shape objects are anchored to a range of text but are free-floating and can be positioned anywhere on the page. Use the ConvertToInlineShape method and the ConvertToShape method to convert shapes from one type to the other. You can convert only pictures, OLE objects, and ActiveX controls to inline shapes. Use the Type property to return the type of inline shape: picture, linked picture, embedded OLE object, linked OLE object, or ActiveX control.

## Properties (36)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified InlineShape object.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified shape.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within an inline shape.
- `LinkFormat As LinkFormat  (read-only)`  
  Returns a LinkFormat object that represents the link options of the specified inline shape that is linked to a file. Read/only.
- `Field As Field  (read-only)`  
  Returns a Field object that represents the field associated with the specified inline shape. Read-only.
- `OLEFormat As OLEFormat  (read-only)`  
  Returns an OLEFormat object that represents the OLE characteristics (other than linking) for the specified inline shape. Read-only.
- `Type As WdInlineShapeType  (read-only)`  
  Returns the type of inline shape. Read-only WdInlineShapeType.
- `Hyperlink As Hyperlink  (read-only)`  
  Returns a Hyperlink object that represents the hyperlink associated with the specified inline shape. Read-only.
- `Height As Single  (read/write)`  
  Returns or sets the height of an inline shape. Read/write Single.
- `Width As Single  (read/write)`  
  Returns or sets the width, in points, of the specified inline shape. Read/write Long.
- `ScaleHeight As Single  (read/write)`  
  Scales the height of the specified inline shape relative to its original size. Read/write Single.
- `ScaleWidth As Single  (read/write)`  
  Scales the width of the specified inline shape relative to its original size. Read/write Single.
- `LockAspectRatio As MsoTriState  (read/write)`  
  MsoTrue if the specified shape retains its original proportions when you resize it. MsoFalse if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.
- `Line As LineFormat  (read-only)`  
  Returns a LineFormat object that contains line formatting properties for the specified shape. Read-only.
- `Fill As FillFormat  (read-only)`  
  Returns a FillFormat object that contains fill formatting properties for the specified shape. Read-only.
- `PictureFormat As PictureFormat  (read/write)`  
  Returns a PictureFormat object that contains picture formatting properties for the inline shape. Read-only.
- `HorizontalLineFormat As HorizontalLineFormat  (read-only)`  
  Returns a HorizontalLineFormat object that contains the horizontal line formatting for the specified InlineShape object. Read-only.
- `Script As Script  (read-only)`  
  Returns a Script object, which represents a block of script or code associated with an image on the specified Web page.
- `TextEffect As TextEffectFormat  (read/write)`  
  Returns a TextEffectFormat object that contains text-effect formatting properties for the specified inline shape. Read-only.
- `AlternativeText As String  (read/write)`  
  Returns or sets a String that represents the alternative text associated with a shape in a webpage. Read/write.
- `IsPictureBullet As Boolean  (read-only)`  
  True indicates that an InlineShape object is a picture bullet. Read-only Boolean.
- `GroupItems As GroupShapes  (read-only)`  
  Returns a GroupShapes collection that represents the shapes that are grouped together for an inline shape. Read-only.
- `HasChart As MsoTriState  (read-only)`  
  True if the specified shape is a chart. Read-only.
- `Chart As Chart  (read-only)`  
  Returns a Chart object that represents a chart within the collection of inline shapes in a document. Read-only.
- `SoftEdge As SoftEdgeFormat  (read-only)`  
  Returns a SoftEdgeFormat object that represents the soft edge formatting for a shape. Read-only.
- `Glow As GlowFormat  (read-only)`  
  Returns a GlowFormat object that represents the formatting properties for a glow effect. Read-only.
- `Reflection As ReflectionFormat  (read-only)`  
  Returns a ReflectionFormat object that represents the reflection formatting for a shape. Read-only.
- `Shadow As ShadowFormat  (read-only)`  
  Returns a ShadowFormat object that represents the shadow formatting for the specified shape. Read-only.
- `HasSmartArt As MsoTriState  (read-only)`  
  Returns True if there is a SmartArt diagram present on the shape. Read-only.
- `SmartArt As SmartArt  (read-only)`  
  Returns a SmartArt object that provides a way to work with the SmartArt associated with the specified inline shape. Read-only.
- `Title As String  (read/write)`  
  Returns or sets a String that contains a title for the specified inline shape. Read/write.
- `GraphicStyle As MsoGraphicStyleIndex  (read/write)`
- `Model3D As Model3DFormat  (read-only)`
- `Decorative As MsoTriState  (read/write)`

## Methods (4)

- `Reset()`  
  Removes changes that were made to an inline shape.
- `Delete()`  
  Deletes the specified inline shape.
- `Select()`  
  Selects the specified inline shape.
- `ConvertToShape() As Shape`  
  Converts an inline shape to a free-floating shape. Returns a Shape object that represents the new shape.
