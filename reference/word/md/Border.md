# Border

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002093B-0000-0000-C000-000000000046}  

Represents a border of an object. The Border object is a member of the Borders collection.

**Remarks:** Use Borders (index), where index identifies the border, to return a single Border object. Index can be one of the WdBorderType constants. Some of the WdBorderType constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. Use the LineStyle property to apply a border line to a Border object. The following example applies a double-line border below the first paragraph in the active document. The following example applies a single-line border around the first character in the selection. The following example adds an art border around each page in the first section. Border objects cannot be added to the Borders collection. The number of members in the Borders collection is finite and varies depending on the type of object. For example, a table has six elements in the Borders collection, whereas a paragraph has four.

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Border object.
- `Visible As Boolean  (read/write)`  
  True if the specified object is visible. Read/write Boolean.
- `ColorIndex As WdColorIndex  (read/write)`  
  Returns or sets the color for the specified border or font object. Read/write WdColorIndex.
- `Inside As Boolean  (read-only)`  
  True if an inside border can be applied to the specified object. Read-only Boolean.
- `LineStyle As WdLineStyle  (read/write)`  
  Returns or sets the border line style for the specified object. Read/write WdLineStyle.
- `LineWidth As WdLineWidth  (read/write)`  
  Returns or sets the line width of an object's border. Read/write.
- `ArtStyle As WdPageBorderArt  (read/write)`  
  Returns or sets the graphical page-border design for a document. Read/write WdPageBorderArt.
- `ArtWidth As Long  (read/write)`  
  Returns or sets the width (in points) of the specified graphical page border. Read/write Long.
- `Color As WdColor  (read/write)`  
  Returns or sets the 24-bit color for the specified Border object.
