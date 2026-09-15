# PageNumbers

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020986-0000-0000-C000-000000000046}  

A collection of PageNumber objects that represent the page numbers in a single header or footer.

**Remarks:** Use the PageNumbers property to return the PageNumbers collection. The following example starts page numbering at 3 for the first section in the active document. Use the Add method to add page numbers to a header or footer. The following example adds a page number to the primary footer in the first section. To add or change page numbers in a document with multiple sections, modify the page numbers in each section or set the LinkToPrevious property to True. Use PageNumbers (index), where index is the index number, to return a single PageNumber object. In most cases, a header or footer contains only one page number, which is index number 1. The following example centers the first page number in the primary header in the first section.

## Properties (13)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified PageNumbers object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of page numbers in the collection. Read-only.
- `NumberStyle As WdPageNumberStyle  (read/write)`  
  Returns or sets a WdPageNumberStyle constant that represents the number style. Read/write.
- `IncludeChapterNumber As Boolean  (read/write)`  
  True if a chapter number is included with page numbers or a caption label. Read/write Boolean.
- `HeadingLevelForChapter As Long  (read/write)`  
  Returns or sets the heading level style that's applied to the chapter titles in the document. Read/write Long.
- `ChapterPageSeparator As WdSeparatorType  (read/write)`  
  Returns or sets the separator character used between the chapter number and the page number. Read/write WdSeparatorType.
- `RestartNumberingAtSection As Boolean  (read/write)`  
  True if page numbering starts at 1 again at the beginning of the specified section. Read/write Boolean.
- `StartingNumber As Long  (read/write)`  
  Returns or sets the starting note number, line number, or page number. Read/write Long.
- `ShowFirstPageNumber As Boolean  (read/write)`  
  True if the page number appears on the first page in the section. Read/write Boolean.
- `DoubleQuote As Boolean  (read/write)`  
  True if Microsoft Word encloses the specified PageNumbers object in double quotation marks ("). Read/write Boolean.

## Methods (2)

- `Item(Index As Long) As PageNumber`  
  Returns an individual PageNumber object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add([PageNumberAlignment As Variant], [FirstPage As Variant]) As PageNumber`  
  Returns a PageNumber object that represents page numbers added to a header or footer in a section.
    - `PageNumberAlignment As Variant` (optional): Can be any WdPageNumberAlignment constant.
    - `FirstPage As Variant` (optional): False to make the first-page header and the first-page footer different from the headers and footers on all subsequent pages in the document. If FirstPage is set to False, a page number isn't added to the first page. If this argument is omitted, the setting is controlled by the DifferentFirstPageHeaderFooter property.
