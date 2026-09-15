# Window

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020962-0000-0000-C000-000000000046}  

Represents a window. Many document characteristics, such as scroll bars and rulers, are actually properties of the window.

**Remarks:** The Window object is a member of the Windows collection. The Windows collection for the Application object contains all the windows in the application, whereas the Windows collection for the Document object contains only the windows that display the specified document. Use Windows (Index), where Index is the window name or the index number, to return a single Window object. The following example maximizes the Document1 window. The index number is the number to the left of the window name on the Window menu. The following example displays the caption of the first window in the Windows collection. Use the Add method or the NewWindow method to add a new window to the Windows collection. Each of the following statements creates a new window for the document in the active window. A colon (:) and a number appear in the window caption when more than one window is open for a document. When you switch the view to print preview, a new window is created. This window is removed from the Windows collection when you close print preview.

## Properties (41)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft OfficeWord application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Window object.
- `ActivePane As Pane  (read-only)`  
  Returns a Pane object that represents the active pane for the specified window. Read-only.
- `Document As Document  (read-only)`  
  Returns a Document object associated with the specified pane, window, or selection. Read-only.
- `Panes As Panes  (read-only)`  
  Returns a Panes collection that represents all the window panes for the specified window.
- `Selection As Selection  (read-only)`  
  Returns the Selection object that represents a selected range or the insertion point. Read-only.
- `Left As Long  (read/write)`  
  Returns or sets a Long that represents the horizontal position of the specified window, measured in points. Read/write.
- `Top As Long  (read/write)`  
  Returns or sets the vertical position of the specified document window, in points. Read/write Long.
- `Width As Long  (read/write)`  
  Returns or sets the width of the specified document window, in points. Read/write Long.
- `Height As Long  (read/write)`  
  Returns or sets the height of the window (in points). Read/write Long.
- `Split As Boolean  (read/write)`  
  True if the window is split into multiple panes. Read/write Boolean.
- `SplitVertical As Long  (read/write)`  
  Returns or sets the vertical split percentage for the specified window. Read/write Long.
- `Caption As String  (read/write)`  
  Returns or sets the caption text for the window that is displayed in the title bar of the document or application window. Read/write String.
- `WindowState As WdWindowState  (read/write)`  
  Returns or sets the state of the specified document window or task window. Read/write WdWindowState.
- `DisplayRulers As Boolean  (read/write)`  
  True if rulers are displayed for the specified window or pane. Read/write Boolean.
- `DisplayVerticalRuler As Boolean  (read/write)`  
  True if a vertical ruler is displayed for the specified window or pane. Read/write Boolean.
- `View As View  (read-only)`  
  Returns a View object that represents the view for the specified window or pane.
- `Type As WdWindowType  (read-only)`  
  Returns the window type. Read-only WdWindowType.
- `Next As Window  (read-only)`  
  Returns the next document window in the collection of open document windows. Read-only.
- `Previous As Window  (read-only)`  
  Returns the previous document window in the collection open document windows. Read-only.
- `WindowNumber As Long  (read-only)`
- `DisplayVerticalScrollBar As Boolean  (read/write)`  
  True if a vertical scroll bar is displayed for the specified window. Read/write Boolean.
- `DisplayHorizontalScrollBar As Boolean  (read/write)`  
  True if a horizontal scroll bar is displayed for the specified window. Read/write Boolean.
- `StyleAreaWidth As Single  (read/write)`  
  Returns or sets the width of the style area in points. Read/write Single.
- `DisplayScreenTips As Boolean  (read/write)`  
  True if comments, footnotes, endnotes, and hyperlinks are displayed as tips. Read/write Boolean.
- `HorizontalPercentScrolled As Long  (read/write)`  
  Returns or sets the horizontal scroll position as a percentage of the document width. Read/write Long.
- `VerticalPercentScrolled As Long  (read/write)`  
  Returns or sets the vertical scroll position as a percentage of the document length. Read/write Long.
- `DocumentMap As Boolean  (read/write)`  
  True if the document map is visible. Read/write Boolean.
- `Active As Boolean  (read-only)`  
  True if the specified window is active. Read-only Boolean.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `IMEMode As WdIMEMode  (read/write)`  
  Returns or sets the default start-up mode for the Japanese Input Method Editor (IME). Read/write WdIMEMode.
- `UsableWidth As Long  (read-only)`  
  Returns the width (in points) of the active working area in the specified document window. Read-only Long.
- `UsableHeight As Long  (read-only)`  
  Returns the height (in points) of the active working area in the specified document window. Read-only Long. .
- `EnvelopeVisible As Boolean  (read/write)`  
  True if the email message header is visible in the document window. The default value is False. Read/write Boolean.
- `DisplayRightRuler As Boolean  (read/write)`  
  True if the vertical ruler appears on the right side of the document window in print layout view. Read/write Boolean.
- `DisplayLeftScrollBar As Boolean  (read/write)`  
  True if the vertical scroll bar appears on the left side of the document window. Read/write Boolean.
- `Visible As Boolean  (read/write)`  
  True if the specified object is visible. Read/write Boolean.
- `Thumbnails As Boolean  (read/write)`  
  Sets or returns a Boolean that represents whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.
- `ShowSourceDocuments As WdShowSourceDocuments  (read/write)`  
  Returns or sets a WdShowSourceDocuments constant that represents how Microsoft Word displays source documents after a compare and merge process. Read/write.
- `Hwnd As Long  (read-only)`  
  Returns a Long that indicates the window handle of the specified window. Read-only.

## Methods (12)

- `Activate()`  
  Activates the specified window.
- `Close([SaveChanges As Variant], [RouteDocument As Variant])`  
  Closes the specified window.
    - `SaveChanges As Variant` (optional): Specifies the save action for the document. Can be one of the following WdSaveOptions constants: wdDoNotSaveChanges, wdPromptToSaveChanges, or wdSaveChanges.
    - `RouteDocument As Variant` (optional): True to route the document to the next recipient. If the document doesn't have a routing slip attached, this argument is ignored.
- `LargeScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant])`  
  Scrolls a window or pane by the specified number of screens.
    - `Down As Variant` (optional): The number of screens to scroll the window down.
    - `Up As Variant` (optional): The number of screens to scroll the window up.
    - `ToRight As Variant` (optional): The number of screens to scroll the window to the right.
    - `ToLeft As Variant` (optional): The number of screens to scroll the window to the left.
- `SmallScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant])`  
  Scrolls a window or pane by the specified number of lines.
    - `Down As Variant` (optional): The number of lines to scroll the window down. A "line" corresponds to the distance scrolled by clicking the down scroll arrow on the vertical scroll bar once.
    - `Up As Variant` (optional): The number of lines to scroll the window up. A "line" corresponds to the distance scrolled by clicking the up scroll arrow on the vertical scroll bar once.
    - `ToRight As Variant` (optional): The number of lines to scroll the window to the right. A "line" corresponds to the distance scrolled by clicking the right scroll arrow on the horizontal scroll bar once.
    - `ToLeft As Variant` (optional): The number of lines to scroll the window to the left. A "line" corresponds to the distance scrolled by clicking the left scroll arrow on the horizontal scroll bar once.
- `NewWindow() As Window`  
  Opens a new window with the same document as the specified window. Returns a Window object.
- `PageScroll([Down As Variant], [Up As Variant])`  
  Scrolls through the specified pane or window page by page.
    - `Down As Variant` (optional): The number of pages to be scrolled down. If this argument is omitted, this value is assumed to be 1.
    - `Up As Variant` (optional): The number of pages to be scrolled up.
- `SetFocus()`  
  Sets the focus of the specified document window to the body of an email message.
- `RangeFromPoint(x As Long, y As Long) As Object`  
  Returns the Range or Shape object that is located at the point specified by the screen position coordinate pair.
    - `x As Long` (required): The horizontal distance (in pixels) from the left edge of the screen to the point.
    - `y As Long` (required): The vertical distance (in pixels) from the top of the screen to the point.
- `ScrollIntoView(obj As Object, [Start As Variant])`  
  Scrolls through the document window so the specified range or shape is displayed in the document window.
    - `obj As Object` (required): A Range or Shape object.
    - `Start As Variant` (optional): True if the upper-left corner of the range or shape appears at the upper-left corner of the document window. False if the lower-right corner of the range or shape appears at the lower-right corner of the document window. The default value is True.
- `GetPoint(ScreenPixelsLeft As Long, ScreenPixelsTop As Long, ScreenPixelsWidth As Long, ScreenPixelsHeight As Long, obj As Object)`  
  Returns the screen coordinates of the specified range or shape.
    - `ScreenPixelsLeft As Long` (required): The variable name to which you want Microsoft Word to return the value for the left edge of the object.
    - `ScreenPixelsTop As Long` (required): The variable name to which you want Word to return the value for the top edge of the object.
    - `ScreenPixelsWidth As Long` (required): The variable name to which you want Word to return the value for the width of the object.
    - `ScreenPixelsHeight As Long` (required): The variable name to which you want Word to return the value for the height of the object.
    - `obj As Object` (required): A Range or Shape object.
- `PrintOut([Background As Variant], [Append As Variant], [Range As Variant], [OutputFileName As Variant], [From As Variant], [To As Variant], [Item As Variant], [Copies As Variant], [Pages As Variant], [PageType As Variant], [PrintToFile As Variant], [Collate As Variant], [ActivePrinterMacGX As Variant], [ManualDuplexPrint As Variant], [PrintZoomColumn As Variant], [PrintZoomRow As Variant], [PrintZoomPaperWidth As Variant], [PrintZoomPaperHeight As Variant])`  
  Prints all or part of the document displayed in the specified window.
    - `Background As Variant` (optional): Set to True to have the macro continue while Microsoft Word prints the document.
    - `Append As Variant` (optional): Set to True to append the specified document to the file name specified by the OutputFileName argument. False to overwrite the contents of OutputFileName.
    - `Range As Variant` (optional): The page range. Can be any WdPrintOutRange constant.
    - `OutputFileName As Variant` (optional): If PrintToFile is True, this argument specifies the path and file name of the output file.
    - `From As Variant` (optional): The starting page number when Range is set to wdPrintFromTo.
    - `To As Variant` (optional): The ending page number when Range is set to wdPrintFromTo.
    - `Item As Variant` (optional): The item to be printed. Can be any WdPrintOutItem constant.
    - `Copies As Variant` (optional): The number of copies to be printed.
    - `Pages As Variant` (optional): The page numbers and page ranges to be printed, separated by commas. For example, "2, 6-10" prints page 2 and pages 6 through 10.
    - `PageType As Variant` (optional): The type of pages to be printed. Can be any WdPrintOutPages constant.
    - `PrintToFile As Variant` (optional): True to send printer instructions to a file. Make sure to specify a file name with OutputFileName.
    - `Collate As Variant` (optional): When printing multiple copies of a document, True to print all pages of the document before printing the next copy.
    - `ActivePrinterMacGX As Variant` (optional): This argument is available only in Microsoft Office Macintosh Edition. For additional information about this argument, consult the language reference Help included with Microsoft Office Macintosh Edition.
    - `ManualDuplexPrint As Variant` (optional): True to print a two-sided document on a printer without a duplex printing kit. If this argument is True, the PrintBackground and PrintReverse properties are ignored. Use the PrintOddPagesInAscendingOrder and PrintEvenPagesInAscendingOrder properties to control the output during manual duplex printing. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `PrintZoomColumn As Variant` (optional): The number of pages you want Word to fit horizontally on one page. Can be 1, 2, 3, or 4. Use with the PrintZoomRow argument to print multiple pages on a single sheet.
    - `PrintZoomRow As Variant` (optional): The number of pages you want Word to fit vertically on one page. Can be 1, 2, or 4. Use with the PrintZoomColumn argument to print multiple pages on a single sheet.
    - `PrintZoomPaperWidth As Variant` (optional): The width to which you want Word to scale printed pages, in twips (20 twips = 1 point; 72 points = 1 inch).
    - `PrintZoomPaperHeight As Variant` (optional): The height to which you want Word to scale printed pages, in twips (20 twips = 1 point; 72 points = 1 inch).
- `ToggleRibbon()`  
  Shows or hides the ribbon.
