# TableOfContents

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020913-0000-0000-C000-000000000046}  

Represents a single table of contents in a document. The TableOfContents object is a member of the TablesOfContents collection. The TablesOfContents collection includes all the tables of contents in a document.

**Remarks:** Use TablesOfContents (Index), where Index is the index number, to return a single TableOfContents object. The index number represents the position of the table of contents in the document. The following example updates the page numbers of the items in the first table of figures in the active document. Use the Add method to add a table of contents to a document. The following example adds a table of contents at the beginning of the active document. The example builds the table of contents from all paragraphs styled as either Heading 1, Heading 2, or Heading 3.

## Properties (15)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TableOfContents object.
- `UseHeadingStyles As Boolean  (read/write)`  
  True if built-in heading styles are used to create a table of contents. Read/write Boolean.
- `UseFields As Boolean  (read/write)`  
  True if Table of Contents Entry (TC) fields are used to create a table of contents or a table of figures. Read/write Boolean.
- `UpperHeadingLevel As Long  (read/write)`  
  Returns or sets the starting heading level for a table of contents. Read/write Long.
- `LowerHeadingLevel As Long  (read/write)`  
  Returns or sets the ending heading level for a table of contents or table of figures. Read/write Long.
- `TableID As String  (read/write)`  
  Returns or sets a one-letter identifier that's used to build a table of contents from TOC fields. Read/write String.
- `HeadingStyles As HeadingStyles  (read-only)`  
  Returns a HeadingStyles object that represents additional styles used to compile a table of contents or table of figures (styles other than the Heading 1 - Heading 9 styles). Read-only.
- `RightAlignPageNumbers As Boolean  (read/write)`  
  True if page numbers are aligned with the right margin in a table of contents. Read/write Boolean.
- `IncludePageNumbers As Boolean  (read/write)`  
  True if page numbers are included in the table of contents. Read/write Boolean.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within the specified table of contents.
- `TabLeader As WdTabLeader  (read/write)`  
  Returns or sets the character between entries and their page numbers in an index, table of authorities, table of contents, or table of figures. Read/write WdTabLeader.
- `UseHyperlinks As Boolean  (read/write)`  
  Returns or sets whether entries in a table of contents should be formatted as hyperlinks when publishing to the Web. Read/write Boolean.
- `HidePageNumbersInWeb As Boolean  (read/write)`  
  Returns or sets whether page numbers in a table of contents or a table of figures should be hidden when publishing to the Web. Read/write Boolean.

## Methods (3)

- `Delete()`  
  Deletes the specified table of contents.
- `UpdatePageNumbers()`  
  Updates the page numbers for items in the specified table of contents.
- `Update()`  
  Updates the entries shown in a table of contents.
