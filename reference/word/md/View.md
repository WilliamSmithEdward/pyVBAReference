# View

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209A5-0000-0000-C000-000000000046}  

Contains the view attributes (such as show all, field shading, and table gridlines) for a window or pane.

**Remarks:** Use the View property to return the View object. The following example sets view options for the active window. Use the Type property to change the view. The following example switches the active window to normal view. Use the Percentage property to change the size of the text on-screen. The following example enlarges the on-screen text to 120 percent. Use the SeekView property to view comments, endnotes, footnotes, or the document header or footer. The following example displays the current footer in the active window in print layout view.

## Properties (56)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified View object.
- `Type As WdViewType  (read/write)`  
  Returns or sets the view type. Read/write WdViewType.
- `FullScreen As Boolean  (read/write)`  
  True if the window is in full-screen view. Read/write Boolean.
- `Draft As Boolean  (read/write)`  
  True if all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display. Read/write Boolean.
- `ShowAll As Boolean  (read/write)`  
  True if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed. Read/write Boolean.
- `ShowFieldCodes As Boolean  (read/write)`  
  True if field codes are displayed. Read/write Boolean.
- `MailMergeDataView As Boolean  (read/write)`  
  True if mail merge data is displayed instead of mail merge fields in the specified window. Read/write Boolean.
- `Magnifier As Boolean  (read/write)`  
  True if the pointer is displayed as a magnifying glass in print preview, indicating that the user can click to zoom in on a particular area of the page or zoom out to see an entire page or spread of pages. Read/write Boolean.
- `ShowFirstLineOnly As Boolean  (read/write)`  
  True if only the first line of body text is shown in outline view. Read/write Boolean.
- `ShowFormat As Boolean  (read/write)`  
  True if character formatting is visible in outline view. Read/write Boolean.
- `Zoom As Zoom  (read-only)`  
  Returns a Zoom object that represents the magnification for the specified view.
- `ShowObjectAnchors As Boolean  (read/write)`  
  True if object anchors are displayed next to items that can be positioned in print layout view. Read/write Boolean.
- `ShowTextBoundaries As Boolean  (read/write)`  
  True if dotted lines are displayed around page margins, text columns, objects, and frames in print layout view. Read/write Boolean.
- `ShowHighlight As Boolean  (read/write)`  
  True if highlight formatting is displayed and printed with a document. Read/write Boolean.
- `ShowDrawings As Boolean  (read/write)`  
  True if objects created with the drawing tools are displayed in print layout view. Read/write Boolean.
- `ShowTabs As Boolean  (read/write)`  
  True if tab characters are displayed. Read/write Boolean.
- `ShowSpaces As Boolean  (read/write)`  
  True if space characters are displayed. Read/write Boolean.
- `ShowParagraphs As Boolean  (read/write)`  
  True if paragraph marks are displayed. Read/write Boolean.
- `ShowHyphens As Boolean  (read/write)`  
  True if optional hyphens are displayed. An optional hyphen indicates where to break a word when it falls at the end of a line. Read/write Boolean.
- `ShowHiddenText As Boolean  (read/write)`  
  True if text formatted as hidden text is displayed. Read/write Boolean.
- `WrapToWindow As Boolean  (read/write)`  
  True if lines wrap at the right edge of the document window rather than at the right margin or the right column boundary. Read/write Boolean.
- `ShowPicturePlaceHolders As Boolean  (read/write)`  
  True if blank boxes are displayed as placeholders for pictures. Read/write Boolean.
- `ShowBookmarks As Boolean  (read/write)`  
  True if square brackets are displayed at the beginning and end of each bookmark. Read/write Boolean.
- `FieldShading As WdFieldShading  (read/write)`  
  Returns or sets on-screen shading for fields. Read/write WdFieldShading.
- `TableGridlines As Boolean  (read/write)`  
  True if table gridlines are displayed. Read/write Boolean.
- `ShowMainTextLayer As Boolean  (read/write)`  
  True if the text in the specified document is visible when the header and footer areas are displayed. This property is equivalent to the Show/Hide Document Text button on the Header and Footer toolbar. Read/write Boolean.
- `SeekView As WdSeekView  (read/write)`  
  Returns or sets the document element displayed in print layout view. The property may be set to any of the WdSeekView constants. Read/write WdSeekView.
- `SplitSpecial As WdSpecialPane  (read/write)`  
  Returns or sets the active window pane. Read/write WdSpecialPane.
- `ShowOptionalBreaks As Boolean  (read/write)`  
  True if Microsoft Word displays optional line breaks. Read/write Boolean.
- `DisplayPageBoundaries As Boolean  (read/write)`  
  True to display the top and bottom margins (white space) and the gray area (gray space) between pages in a document. Read/write Boolean.
- `ShowRevisionsAndComments As Boolean  (read/write)`  
  True for Microsoft Word to display revisions and comments that were made to a document with Track Changes enabled. Read/write Boolean.
- `ShowComments As Boolean  (read/write)`  
  True for Microsoft Word to display the comments in a document. Read/write Boolean.
- `ShowInsertionsAndDeletions As Boolean  (read/write)`  
  True for Microsoft Word to display insertions and deletions that were made to a document with Track Changes enabled. Read/write Boolean.
- `ShowFormatChanges As Boolean  (read/write)`  
  True for Microsoft Word to display formatting changes made to a document with Track Changes enabled. Read/write Boolean.
- `RevisionsBalloonWidth As Single  (read/write)`  
  Sets or returns a Single representing the global setting in Microsoft Word that specifies the width of the revision balloons. Read/write.
- `RevisionsBalloonWidthType As WdRevisionsBalloonWidthType  (read/write)`  
  Sets or returns a WdRevisionsBalloonWidthType constant representing the global setting that specifies how Microsoft Word measures the width of revision balloons. Read/write.
- `RevisionsBalloonSide As WdRevisionsBalloonMargin  (read/write)`  
  Sets or returns a WdRevisionsBalloonMargin constant that specifies whether Word displays revision balloons in the left or right margin in a document.
- `RevisionsBalloonShowConnectingLines As Boolean  (read/write)`  
  True for Microsoft Word to display connecting lines from the text to the revision and comment balloons. Read/write Boolean.
- `ReadingLayout As Boolean  (read/write)`  
  Sets or returns a Boolean that represents whether a document is being viewed in reading layout view. .
- `ShowXMLMarkup As Long  (read/write)`  
  Returns a Long that represents whether XML tags are visible in a document.
- `ShadeEditableRanges As Long  (read/write)`  
  Returns or sets a Long that represents whether shading is applied to the ranges in a document for which users have permission to modify. .
- `ShowInkAnnotations As Boolean  (read/write)`  
  Returns or sets Boolean that shows or hides handwritten ink annotations. True displays ink annotations. False hides ink annotations.
- `DisplayBackgrounds As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether background colors and images are shown when a document is displayed in print layout view. .
- `ReadingLayoutActualView As Boolean  (read/write)`  
  Sets or returns a Boolean that represents whether pages displayed in reading layout view are displayed using the same layout as printed pages.
- `ReadingLayoutTruncateMargins As WdReadingLayoutMargin  (read/write)`  
  Returns or sets a WdReadingLayoutMargin constant that represents whether margins are visible or hidden when a document is viewed in Full Screen Reading view. Read/write.
- `Panning As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Microsoft Word is in Panning mode. Read/write.
- `ShowCropMarks As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to show crop marks in the corners of pages to indicate where margins are located. Read/write.
- `MarkupMode As WdRevisionsMode  (read/write)`  
  Returns or sets a WdRevisionsMode constant that represents the display mode for tracked changes. Read/write.
- `ConflictMode As Boolean  (read/write)`  
  True if the document is in conflict mode view. Read/write.
- `ShowOtherAuthors As Boolean  (read/write)`  
  True if other authors' presence should be visible in the document. Read/write Boolean.
- `RevisionsFilter As RevisionsFilter  (read-only)`  
  Returns an instance of a RevisionsFilter object. Read-only.
- `PageColor As WdPageColor  (read/write)`  
  Returns and sets the page color in Reading mode. Read/write WdPageColor.
- `ColumnWidth As WdColumnWidth  (read/write)`  
  Returns or gets a constant that determines the column width in reading mode. Read/write WdColumnWidth
- `PageMovementType As WdPageMovementType  (read/write)`  
  Returns or sets the page movement type. Read/write WdPageMovementType.

## Methods (8)

- `CollapseOutline([Range As Variant])`  
  Collapses the text under the selection or the specified range by one heading level.
    - `Range As Variant` (optional): The range of paragraphs to be collapsed. If this argument is omitted, the entire selection is collapsed.
- `ExpandOutline([Range As Variant])`  
  Expands the text under the selection by one heading level.
    - `Range As Variant` (optional): The range of paragraphs to be expanded. If this argument is omitted, the entire selection is expanded.
- `ShowAllHeadings()`  
  Switches between showing all text (headings and body text) and showing only headings.
- `ShowHeading(Level As Long)`  
  Shows all headings up to the specified heading level and hides subordinate headings and body text.
    - `Level As Long` (required): The outline heading level (a number from 1 to 9).
- `PreviousHeaderFooter()`  
  Moves to the previous header or footer, depending on whether a header or footer is displayed in the view.
- `NextHeaderFooter()`  
  Moves to the next header or footer, depending on whether a header or footer is displayed in the view.
- `ExpandAllHeadings()`  
  Expands all the headings in the document.
- `CollapseAllHeadings()`  
  Collapses all the headings in the document.
