# Pane

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020895-0000-0000-C000-000000000046}  

Represents a pane of a window.

**Remarks:** Pane objects exist only for worksheets and Microsoft Excel 4.0 macro sheets. The Pane object is a member of the Panes collection. The Panes collection contains all of the panes shown in a single window.

**Example:**

```vba
Worksheets(1).Activate
ActiveWindow.Split = True
ActiveWindow.Panes(3).ScrollRow = 5
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `ScrollColumn As Long  (read/write)`  
  Returns or sets the number of the leftmost column in the pane or window. Read/write Long.
- `ScrollRow As Long  (read/write)`  
  Returns or sets the number of the row that appears at the top of the pane or window. Read/write Long.
- `VisibleRange As Range  (read-only)`  
  Returns a Range object that represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range. Read-only.

## Methods (6)

- `Activate() As Boolean`  
  Activates the pane.
- `LargeScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant]) As Variant`  
  Scrolls the contents of the window by pages.
    - `Down As Variant` (optional): The number of pages to scroll the contents down.
    - `Up As Variant` (optional): The number of pages to scroll the contents up.
    - `ToRight As Variant` (optional): The number of pages to scroll the contents to the right.
    - `ToLeft As Variant` (optional): The number of pages to scroll the contents to the left.
- `SmallScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant]) As Variant`  
  Scrolls the contents of the window by rows or columns.
    - `Down As Variant` (optional): The number of rows to scroll the contents down.
    - `Up As Variant` (optional): The number of rows to scroll the contents up.
    - `ToRight As Variant` (optional): The number of columns to scroll the contents to the right.
    - `ToLeft As Variant` (optional): The number of columns to scroll the contents to the left.
- `ScrollIntoView(Left As Long, Top As Long, Width As Long, Height As Long, [Start As Variant])`  
  Scrolls the document window so that the contents of a specified rectangular area are displayed in either the upper-left or lower-right corner of the document window or pane (depending on the value of the _Start_ argument).
    - `Left As Long` (required): The horizontal position of the rectangle (in points) from the left edge of the document window or pane.
    - `Top As Long` (required): The vertical position of the rectangle (in points) from the top of the document window or pane.
    - `Width As Long` (required): The width of the rectangle, in points.
    - `Height As Long` (required): The height of the rectangle, in points.
    - `Start As Variant` (optional): True to have the upper-left corner of the rectangle appear in the upper-left corner of the document window or pane. False to have the lower-right corner of the rectangle appear in the lower-right corner of the document window or pane. The default value is True.
- `PointsToScreenPixelsX(Points As Long) As Long`  
  Returns or sets a pixel point on the screen.
    - `Points As Long` (required): Location of the pixel on the screen.
- `PointsToScreenPixelsY(Points As Long) As Long`  
  Returns or sets the location of the pixel on the screen.
    - `Points As Long` (required): Location of the starting point.
