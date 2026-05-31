# TextFrame

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002443D-0000-0000-C000-000000000046}  

Represents the text frame in a Shape object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.

**Remarks:** Use the TextFrame property of the Shape object to return a TextFrame object.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes.AddShape(msoShapeRectangle, _
 0, 0, 250, 140).TextFrame
 .Characters.Text = "Here is some test text"
 .MarginBottom = 10
 .MarginLeft = 10
 .MarginRight = 10
 .MarginTop = 10
End With
```

## Properties (15)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `MarginBottom As Single  (read/write)`  
  Returns or sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `MarginLeft As Single  (read/write)`  
  Returns or sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `MarginRight As Single  (read/write)`  
  Returns or sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `MarginTop As Single  (read/write)`  
  Returns or sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write Single.
- `Orientation As MsoTextOrientation  (read/write)`  
  Returns or sets a Long value that represents the text frame orientation.
- `HorizontalAlignment As XlHAlign  (read/write)`  
  Returns or sets an XlHAlign value that represents the horizontal alignment for the specified object.
- `VerticalAlignment As XlVAlign  (read/write)`  
  Returns or sets an XlVAlign value that represents the vertical alignment of the specified object.
- `AutoSize As Boolean  (read/write)`  
  True if the size of the specified object is changed automatically to fit text within its boundaries. Read/write Boolean.
- `ReadingOrder As Long  (read/write)`  
  Returns or sets the reading order for the specified object. Can be one of the following XlReadingOrder constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext. Read/write Long.
- `AutoMargins As Boolean  (read/write)`  
  Returns or sets whether Excel automatically calculates text frame margins. Read/write.
- `VerticalOverflow As XlOartVerticalOverflow  (read/write)`  
  Returns or sets the vertical overflow setting for the specified object. Read/write.
- `HorizontalOverflow As XlOartHorizontalOverflow  (read/write)`  
  Returns or sets the horizontal overflow setting for the specified object. Read/write.

## Methods (1)

- `Characters([Start As Variant], [Length As Variant]) As Characters`  
  Returns a Characters object that represents a range of characters within a shape's text frame. Use the Characters object to add and format characters within the text frame.
    - `Start As Variant` (optional): The first character to be returned. If this argument is either set to 1 or omitted, the Characters method returns a range of characters starting with the first character.
    - `Length As Variant` (optional): The number of characters to be returned. If this argument is omitted, the Characters method returns the remainder of the string (everything after the character that was set as the _Start_ argument).
