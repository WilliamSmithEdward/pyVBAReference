# PageSetup

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020971-0000-0000-C000-000000000046}  

Represents the page setup description. The PageSetup object contains all the page setup attributes of a document (such as left margin, bottom margin, and paper size) as properties.

**Remarks:** Use the PageSetup property to return the PageSetup object. The following example sets the first section in the active document to landscape orientation and then prints the document. The following example sets all the margins for the document named "Sales.doc."

## Properties (35)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified PageSetup object.
- `TopMargin As Single  (read/write)`  
  Returns or sets the distance (in points) between the top edge of the page and the top boundary of the body text. Read/write Single.
- `BottomMargin As Single  (read/write)`  
  Returns or sets the distance (in points) between the bottom edge of the page and the bottom boundary of the body text. Read/write Single.
- `LeftMargin As Single  (read/write)`  
  Returns or sets the distance (in points) between the left edge of the page and the left boundary of the body text. Read/write Single.
- `RightMargin As Single  (read/write)`  
  Returns or sets the distance (in points) between the right edge of the page and the right boundary of the body text. Read/write Single.
- `Gutter As Single  (read/write)`  
  Returns or sets the amount (in points) of extra margin space added to each page in a document or section for binding. Read/write Single.
- `PageWidth As Single  (read/write)`  
  Returns or sets the width of the page in points. Read/write Single.
- `PageHeight As Single  (read/write)`  
  Returns or sets the height of the page in points. Read/write Single.
- `Orientation As WdOrientation  (read/write)`  
  Returns or sets the orientation of the page. Read/write WdOrientation.
- `FirstPageTray As WdPaperTray  (read/write)`  
  Returns or sets the paper tray to use for the first page of a document or section. Read/write WdPaperTray.
- `OtherPagesTray As WdPaperTray  (read/write)`  
  Returns or sets the paper tray to be used for all but the first page of a document or section. Read/write WdPaperTray.
- `VerticalAlignment As WdVerticalAlignment  (read/write)`  
  Returns or sets the vertical alignment of text on each page in a document or section. Read/write WdVerticalAlignment.
- `MirrorMargins As Long  (read/write)`  
  True if the inside and outside margins of facing pages are the same width. Read/write Long.
- `HeaderDistance As Single  (read/write)`  
  Returns or sets the distance (in points) between the header and the top of the page. Read/write Single.
- `FooterDistance As Single  (read/write)`  
  Returns or sets the distance (in points) between the footer and the bottom of the page. Read/write Single.
- `SectionStart As WdSectionStart  (read/write)`  
  Returns or sets the type of section break for the specified object. Read/write WdSectionStart.
- `OddAndEvenPagesHeaderFooter As Long  (read/write)`  
  True if the specified PageSetup object has different headers and footers for odd-numbered and even-numbered pages. Read/write Long.
- `DifferentFirstPageHeaderFooter As Long  (read/write)`  
  True if a different header or footer is used on the first page. Can be True, False, or wdUndefined. Read/write Long.
- `SuppressEndnotes As Long  (read/write)`  
  True if endnotes are printed at the end of the next section that doesn't suppress endnotes. Read/write Long.
- `LineNumbering As LineNumbering  (read/write)`  
  Returns or sets a LineNumbering object that represents the line numbers for the specified PageSetup object.
- `TextColumns As TextColumns  (read/write)`  
  Returns a TextColumns collection that represents the set of text columns for the specified PageSetup object.
- `PaperSize As WdPaperSize  (read/write)`  
  Returns or sets the paper size. Read/write WdPaperSize.
- `TwoPagesOnOne As Boolean  (read/write)`  
  True if Microsoft Word prints the specified document two pages per sheet. Read/write Boolean.
- `CharsLine As Single  (read/write)`  
  Returns or sets the number of characters per line in the document grid. Read/write Single.
- `LinesPage As Single  (read/write)`  
  Returns or sets the number of lines per page in the document grid. Read/write Single.
- `ShowGrid As Boolean  (read/write)`  
  True to display the page grid. False to hide it. Read/write Boolean.
- `GutterStyle As WdGutterStyleOld  (read/write)`  
  Returns or sets whether Microsoft Word uses gutters for the current document based on a right-to-left language or a left-to-right language. Read/write WdGutterStyleOld.
- `SectionDirection As WdSectionDirection  (read/write)`  
  Returns or sets the reading order and alignment for the specified sections. Read/write WdSectionDirection.
- `LayoutMode As WdLayoutMode  (read/write)`  
  Returns or sets the layout mode for the current document. Read/write WdLayoutMode.
- `GutterPos As WdGutterStyle  (read/write)`  
  Returns or sets on which side the gutter appears in a document. Read/write WdGutterStyle.
- `BookFoldPrinting As Boolean  (read/write)`  
  True for Microsoft Word to print a document in a series of booklets so that the printed pages can be folded and read as a book. Read/write Boolean.
- `BookFoldRevPrinting As Boolean  (read/write)`  
  True for Microsoft Word to reverse the printing order for book fold printing of bidirectional or Asian language documents. Read/write Boolean.
- `BookFoldPrintingSheets As Long  (read/write)`  
  Returns or sets a Long which represents the number of pages for each booklet. Read/write Boolean.

## Methods (2)

- `TogglePortrait()`  
  Switches between portrait and landscape page orientations for a document or section.
- `SetAsTemplateDefault()`  
  Sets the specified page setup formatting as the default for the active document and all new documents based on the active template.
