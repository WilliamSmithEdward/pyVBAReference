# Pane

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020960-0000-0000-C000-000000000046}  

Represents a window pane. The Pane object is a member of the Panes collection. The Panes collection includes all the window panes for a single window.

**Remarks:** Use Panes (Index), where Index is the index number, to return a single Pane object. The following example closes the active pane. Use the Add method or the Split property to add a window pane. The following example splits the active window at 20 percent of the current window size. The following example splits the active window in half. Use the SplitSpecial property to show comments, footnotes, or endnotes in a separate pane. A window has more than one pane if the window is split or the view is not print layout view and information such as footnotes or comments are displayed. The following example displays the comments pane in normal view and then prompts to close the pane.

## Properties (18)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Pane object.
- `Document As Document  (read-only)`  
  Returns a Document object associated with the specified pane. Read-only.
- `Selection As Selection  (read-only)`  
  Returns the Selection object that represents a selection or the insertion point within a document pane. Read-only.
- `DisplayRulers As Boolean  (read/write)`  
  True if rulers are displayed for the specified pane. Read/write Boolean.
- `DisplayVerticalRuler As Boolean  (read/write)`  
  True if a vertical ruler is displayed for the specified pane. Read/write Boolean.
- `Zooms As Zooms  (read-only)`  
  Returns a Zooms collection that represents the magnification options for each view (such as normal view, outline view or print layout view).
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `View As View  (read-only)`  
  Returns a View object that represents the view for the specified pane.
- `Next As Pane  (read-only)`  
  Returns a Pane object that represents the next document pane in the collection. Read-only.
- `Previous As Pane  (read-only)`  
  Returns a Pane object that represents the previous document pane in the collection. Read-only.
- `HorizontalPercentScrolled As Long  (read/write)`  
  Returns or sets the horizontal scroll position as a percentage of the document width. Read/write Long.
- `VerticalPercentScrolled As Long  (read/write)`  
  Returns or sets the vertical scroll position as a percentage of the document length. Read/write Long.
- `MinimumFontSize As Long  (read/write)`  
  Returns or sets the minimum font size (in points) displayed for the specified pane. Read/write Long.
- `BrowseWidth As Long  (read-only)`  
  Returns the width (in points) of the area in which text wraps in the specified pane. Read-only Long.
- `Frameset As Frameset  (read-only)`  
  Returns a Frameset object that represents an entire frames page or a single frame on a frames page. Read-only.
- `Pages As Pages  (read-only)`  
  Returns a Pages collection that represents the pages in a document.

## Methods (8)

- `Activate()`  
  Activates the specified pane.
- `Close()`  
  Closes the specified Mail Merge data source, pane, or task.
- `LargeScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant])`  
  Scrolls a window or pane by the specified number of screens.
    - `Down As Variant` (optional): The number of screens to scroll the window down.
    - `Up As Variant` (optional): The number of screens to scroll the window up.
    - `ToRight As Variant` (optional): The number of screens to scroll the window to the right.
    - `ToLeft As Variant` (optional): The number of screens to scroll the window to the left.
- `SmallScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant])`  
  Scrolls a window by the specified number of lines.
    - `Down As Variant` (optional): The number of lines to scroll the window down. A "line" corresponds to the distance scrolled by clicking the down scroll arrow on the vertical scroll bar once.
    - `Up As Variant` (optional): The number of lines to scroll the window up. A "line" corresponds to the distance scrolled by clicking the up scroll arrow on the vertical scroll bar once.
    - `ToRight As Variant` (optional): The number of lines to scroll the window to the right. A "line" corresponds to the distance scrolled by clicking the right scroll arrow on the horizontal scroll bar once.
    - `ToLeft As Variant` (optional): The number of lines to scroll the window to the left. A "line" corresponds to the distance scrolled by clicking the left scroll arrow on the horizontal scroll bar once.
- `AutoScroll(Velocity As Long)`  
  Scrolls automatically through the specified pane.
    - `Velocity As Long` (required): The speed for scrolling. Can be a number from -100 through 100. Use -100 for full-speed backward scrolling, and use 100 for full-speed forward scrolling.
- `PageScroll([Down As Variant], [Up As Variant])`  
  Scrolls through the specified window page by page.
    - `Down As Variant` (optional): The number of pages to be scrolled down. If this argument is omitted, this value is assumed to be 1.
    - `Up As Variant` (optional): The number of pages to be scrolled up.
- `NewFrameset()`  
  Creates a new frames page based on the specified pane.
- `TOCInFrameset()`  
  Creates a table of contents based on the specified document and puts it in a new frame on the left side of the frames page.
