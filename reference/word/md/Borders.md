# Borders

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002093C-0000-0000-C000-000000000046}  

A collection of Border objects that represent the borders of an object.

**Remarks:** Use the Borders property to return the Borders collection. The following example applies the default border around the first paragraph in the active document. Border objects cannot be added to the Borders collection. The number of members in the Borders collection is finite and varies depending on the type of object. For example, a table has six elements in the Borders collection, whereas a paragraph has four. Use Borders (index), where index identifies the border, to return a single Border object. Index can be one of the WdBorderType constants. Some of the WdBorderType constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. Use the LineStyle property to apply a border line to a Border object. The following example applies a double-line border below the first paragraph in the active document. The following example applies a single-line border around the first character in the selection. The following example adds an art border around each page in the first section.

## Properties (28)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Borders collection.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of items in the Borders collection. Read-only Long.
- `Enable As Long  (read/write)`  
  Returns or sets border formatting for the specified object. Read/write Long.
- `DistanceFromTop As Long  (read/write)`  
  Returns or sets the space (in points) between the text and the top border. Read/write Long.
- `Shadow As Boolean  (read/write)`  
  True if the specified border is formatted as shadowed. Read/write Boolean.
- `InsideLineStyle As WdLineStyle  (read/write)`  
  Returns or sets the inside border for the specified object. .
- `OutsideLineStyle As WdLineStyle  (read/write)`  
  Returns or sets the outside border for the specified object. .
- `InsideLineWidth As WdLineWidth  (read/write)`  
  Returns or sets the line width of the inside border of an object. .
- `OutsideLineWidth As WdLineWidth  (read/write)`  
  Returns or sets the line width of the outside border of an object. Read/write.
- `InsideColorIndex As WdColorIndex  (read/write)`  
  Returns or sets the color of the inside borders. Read/write WdColorIndex.
- `OutsideColorIndex As WdColorIndex  (read/write)`  
  Returns or sets the color of the outside borders. Read/write WdColorIndex.
- `DistanceFromLeft As Long  (read/write)`  
  Returns or sets the space (in points) between the text and the left border. Read/write Long.
- `DistanceFromBottom As Long  (read/write)`  
  Returns or sets the space (in points) between the text and the bottom border. Read/write Long.
- `DistanceFromRight As Long  (read/write)`  
  Returns or sets the space (in points) between the right edge of the text and the right border. Read/write Long.
- `AlwaysInFront As Boolean  (read/write)`  
  True if page borders are displayed in front of the document text. Read/write Boolean.
- `SurroundHeader As Boolean  (read/write)`  
  True if a page border encompasses the document header. Read/write Boolean.
- `SurroundFooter As Boolean  (read/write)`  
  True if a page border encompasses the document footer. Read/write Boolean.
- `JoinBorders As Boolean  (read/write)`  
  True if vertical borders at the edges of paragraphs and tables are removed so that the horizontal borders can connect to the page border. Read/write Boolean.
- `HasHorizontal As Boolean  (read-only)`  
  True if a horizontal border can be applied to the object. Read-only Boolean.
- `HasVertical As Boolean  (read-only)`  
  True if a vertical border can be applied to the specified object. Read-only Boolean.
- `DistanceFrom As WdBorderDistanceFrom  (read/write)`  
  Returns or sets a value that indicates whether the specified page border is measured from the edge of the page or from the text it surrounds. Read/write WdBorderDistanceFrom.
- `EnableFirstPageInSection As Boolean  (read/write)`  
  True if page borders are enabled for the first page in the section. Read/write Boolean.
- `EnableOtherPagesInSection As Boolean  (read/write)`  
  True if page borders are enabled for all pages in the section except for the first page. Read/write Boolean.
- `InsideColor As WdColor  (read/write)`  
  Returns or sets the 24-bit color of the inside borders. Read/write.
- `OutsideColor As WdColor  (read/write)`  
  Returns or sets the 24-bit color of the outside borders. .

## Methods (2)

- `Item(Index As WdBorderType) As Border`  
  Returns a border in a range or selection.
    - `Index As WdBorderType` (required): The border to be returned.
- `ApplyPageBordersToAllSections()`  
  Applies the specified page-border formatting to all sections in a document.
