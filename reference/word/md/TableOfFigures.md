# TableOfFigures

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020921-0000-0000-C000-000000000046}  

Represents a single table of figures in a document. The TableOfFigures object is a member of the TablesOfFigures collection. The TablesOfFigures collection includes all the tables of figures in a document.

**Remarks:** Use TablesOfFigures (Index), where Index is the index number, to return a single TableOfFigures object. The index number represents the position of the table of figures in the document. The following example updates the page numbers of the items in the first table of figures in the active document. Use the Add method to add a table of figures to a document. A table of figures lists figure captions in the order in which they appear in the document. The following example replaces the selection in the active document with a table of figures that includes caption labels and page numbers.

## Properties (17)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TableOfFigures object.
- `Caption As String  (read/write)`  
  Returns or sets the label that identifies the items to be included in a table of figures. Read/write String.
- `IncludeLabel As Boolean  (read/write)`  
  True if the caption label and caption number are included in a table of figures. Read/write Boolean.
- `RightAlignPageNumbers As Boolean  (read/write)`  
  True if page numbers are aligned with the right margin in an table of figures. Read/write Boolean.
- `UseHeadingStyles As Boolean  (read/write)`  
  True if built-in heading styles are used to create a table of figures. Read/write Boolean.
- `LowerHeadingLevel As Long  (read/write)`  
  Returns or sets the ending heading level for a table of figures. Read/write Long.
- `UpperHeadingLevel As Long  (read/write)`  
  Returns or sets the starting heading level for a table of figures. Read/write Long.
- `IncludePageNumbers As Boolean  (read/write)`  
  True if page numbers are included in the table of figures. Read/write Boolean.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within the specified table of figures.
- `UseFields As Boolean  (read/write)`  
  True if Table of Contents Entry (TC) fields are used to create a table of figures. Read/write Boolean.
- `TableID As String  (read/write)`  
  Returns or sets a one-letter identifier that is used to build a table of figures from TOC fields. Read/write String.
- `HeadingStyles As HeadingStyles  (read-only)`  
  Returns a HeadingStyles object that represents additional styles used to compile a table of contents or table of figures (styles other than the Heading 1 - Heading 9 styles). Read-only.
- `TabLeader As WdTabLeader  (read/write)`  
  Returns or sets the character between entries and their page numbers in an table of figures. Read/write WdTabLeader.
- `UseHyperlinks As Boolean  (read/write)`  
  Returns or sets whether entries in a table of figures should be formatted as hyperlinks when publishing to the Web. Read/write Boolean.
- `HidePageNumbersInWeb As Boolean  (read/write)`  
  Returns or sets whether page numbers in a table of contents or a table of figures should be hidden when publishing to the Web. Read/write Boolean.

## Methods (3)

- `Delete()`  
  Deletes the specified table of figures.
- `UpdatePageNumbers()`  
  Updates the page numbers for items in a table of figures.
- `Update()`  
  Updates the entries shown in a table of figures.
