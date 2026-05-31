# Window

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020893-0000-0000-C000-000000000046}  

Represents a window.

**Remarks:** Many worksheet characteristics, such as scroll bars and gridlines, are actually properties of the window. The Window object is a member of the Windows collection. The Windows collection for the Application object contains all the windows in the application, whereas the Windows collection for the Workbook object contains only the windows in the specified workbook.

**Example:**

```vba
Windows(1).WindowState = xlMaximized
```

## Properties (55)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ActiveCell As Range  (read-only)`  
  Returns a Range object that represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails. Read-only.
- `ActiveChart As Chart  (read-only)`  
  Returns a Chart object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns Nothing.
- `ActivePane As Pane  (read-only)`  
  Returns a Pane object that represents the active pane in the window. Read-only.
- `ActiveSheet As Object  (read-only)`  
  Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns Nothing if no sheet is active.
- `Caption As Variant  (read/write)`  
  Returns or sets a Variant value that represents the name that appears in the title bar of the document window.
- `DisplayFormulas As Boolean  (read/write)`  
  True if the window is displaying formulas; False if the window is displaying values. Read/write Boolean.
- `DisplayGridlines As Boolean  (read/write)`  
  True if gridlines are displayed. Read/write Boolean.
- `DisplayHeadings As Boolean  (read/write)`  
  True if both row and column headings are displayed; False if no headings are displayed. Read/write Boolean.
- `DisplayHorizontalScrollBar As Boolean  (read/write)`  
  True if the horizontal scroll bar is displayed. Read/write Boolean.
- `DisplayOutline As Boolean  (read/write)`  
  True if outline symbols are displayed. Read/write Boolean.
- `DisplayVerticalScrollBar As Boolean  (read/write)`  
  True if the vertical scroll bar is displayed. Read/write Boolean.
- `DisplayWorkbookTabs As Boolean  (read/write)`  
  True if the workbook tabs are displayed. Read/write Boolean.
- `DisplayZeros As Boolean  (read/write)`  
  True if zero values are displayed. Read/write Boolean.
- `EnableResize As Boolean  (read/write)`  
  True if the window can be resized. Read/write Boolean.
- `FreezePanes As Boolean  (read/write)`  
  True if split panes are frozen. Read/write Boolean.
- `GridlineColor As Long  (read/write)`  
  Returns or sets the gridline color as an RGB value. Read/write Long.
- `GridlineColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the gridline color as an index into the current color palette or as an XlColorIndex constant.
- `Height As Double  (read/write)`  
  Returns or sets a Double value that represents the height, in points, of the window.
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `Left As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the left edge of the client area to the left edge of the window.
- `OnWindow As String  (read/write)`  
  Returns or sets the name of the procedure that's run whenever you activate a window. Read/write String.
- `Panes As Panes  (read-only)`  
  Returns a Panes collection that represents all the panes in the specified window. Read-only.
- `RangeSelection As Range  (read-only)`  
  Returns a Range object that represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet. Read-only.
- `ScrollColumn As Long  (read/write)`  
  Returns or sets the number of the leftmost column in the pane or window. Read/write Long.
- `ScrollRow As Long  (read/write)`  
  Returns or sets the number of the row that appears at the top of the pane or window. Read/write Long.
- `SelectedSheets As Sheets  (read-only)`  
  Returns a Sheets collection that represents all the selected sheets in the specified window. Read-only.
- `Selection As Object  (read-only)`  
  Returns the specified window, for a Windows object.
- `Split As Boolean  (read/write)`  
  True if the window is split. Read/write Boolean.
- `SplitColumn As Long  (read/write)`  
  Returns or sets the column number where the window is split into panes (the number of columns to the left of the split line). Read/write Long.
- `SplitHorizontal As Double  (read/write)`  
  Returns or sets the location of the horizontal window split, in points. Read/write Double.
- `SplitRow As Long  (read/write)`  
  Returns or sets the row number where the window is split into panes (the number of rows above the split). Read/write Long.
- `SplitVertical As Double  (read/write)`  
  Returns or sets the location of the vertical window split, in points. Read/write Double.
- `TabRatio As Double  (read/write)`  
  Returns or sets the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar (as a number between 0 (zero) and 1; the default value is 0.6). Read/write Double.
- `Top As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the top edge of the window to the top edge of the usable area (below the menus, any toolbars docked at the top, and the formula bar).
- `Type As XlWindowType  (read-only)`  
  Returns or sets an XlWindowType value that represents the window type.
- `UsableHeight As Double  (read-only)`  
  Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-only Double.
- `UsableWidth As Double  (read-only)`  
  Returns the maximum width of the space that a window can occupy in the application window area, in points. Read-only Double.
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.
- `VisibleRange As Range  (read-only)`  
  Returns a Range object that represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range. Read-only.
- `Width As Double  (read/write)`  
  Returns or sets a Double value that represents the width, in points, of the window.
- `WindowNumber As Long  (read-only)`  
  Returns the window number. For example, a window named Book1.xls:2 has 2 as its window number. Most windows have the window number 1. Read-only Long.
- `WindowState As XlWindowState  (read/write)`  
  Returns or sets the state of the window. Read/write XlWindowState.
- `Zoom As Variant  (read/write)`  
  Returns or sets a Variant value that represents the display size of the window, as a percentage (100 equals normal size, 200 equals double size, and so on).
- `View As XlWindowView  (read/write)`  
  Returns or sets the view showing in the window. Read/write XlWindowView.
- `DisplayRightToLeft As Boolean  (read/write)`  
  True if the specified window is displayed from right to left instead of from left to right. False if the object is displayed from left to right. Read-only Boolean.
- `SheetViews As SheetViews  (read-only)`  
  Returns the SheetViews object for the specified window. Read-only.
- `ActiveSheetView As Object  (read-only)`  
  Returns an object that represents the view of the active sheet in the specified window. Read-only.
- `DisplayRuler As Boolean  (read/write)`  
  True if a ruler is displayed for the specified window. Read/write Boolean.
- `AutoFilterDateGrouping As Boolean  (read/write)`  
  True if the auto filter for date grouping is currently displayed in the specified window. Read/write Boolean.
- `DisplayWhitespace As Boolean  (read/write)`  
  True if whitespace is displayed. Read/write Boolean.
- `Hwnd As Long  (read-only)`  
  Returns a Long that indicates the window handle of the specified window. Read-only.
- `DisplayDataTypeIcons As Boolean  (read/write)`

## Methods (14)

- `Activate() As Variant`  
  Brings the window to the front of the z-order.
- `ActivateNext() As Variant`  
  Activates the specified window and then sends it to the back of the window z-order.
- `ActivatePrevious() As Variant`  
  Activates the specified window and then activates the window at the back of the window z-order.
- `Close([SaveChanges As Variant], [Filename As Variant], [RouteWorkbook As Variant]) As Boolean`  
  Closes the object.
    - `SaveChanges As Variant` (optional): If there are no changes to the workbook, this argument is ignored. If there are changes to the workbook and the workbook appears in other open windows, this argument is ignored. If there are changes to the workbook but the workbook doesn't appear in any other open windows, this argument specifies whether changes should be saved.If set to True, changes are saved to the workbook. If there is not yet a file name associated with the workbook, _FileName_ is used. If _FileName_ is omitted, the user is asked to supply a file name.
    - `Filename As Variant` (optional): Save changes under this file name.
    - `RouteWorkbook As Variant` (optional): If the workbook doesn't need to be routed to the next recipient (if it has no routing slip or has already been routed), this argument is ignored. Otherwise, Microsoft Excel routes the workbook according to the value of this parameter. If set to True, the workbook is sent to the next recipient. If set to False, the workbook is not sent. If omitted, the user is asked whether the workbook should be sent.
- `LargeScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant]) As Variant`  
  Scrolls the contents of the window by pages.
    - `Down As Variant` (optional): The number of pages to scroll the contents down.
    - `Up As Variant` (optional): The number of pages to scroll the contents up.
    - `ToRight As Variant` (optional): The number of pages to scroll the contents to the right.
    - `ToLeft As Variant` (optional): The number of pages to scroll the contents to the left.
- `NewWindow() As Window`  
  Creates a new window or a copy of the specified window.
- `PrintPreview([EnableChanges As Variant]) As Variant`  
  Shows a preview of the object as it would look when printed.
    - `EnableChanges As Variant` (optional): Pass a Boolean value to specify if the user can change the margins and other page setup options available in print preview.
- `ScrollWorkbookTabs([Sheets As Variant], [Position As Variant]) As Variant`  
  Scrolls through the workbook tabs at the bottom of the window. Doesn't affect the active sheet in the workbook.
    - `Sheets As Variant` (optional): The number of sheets to scroll by. Use a positive number to scroll forward, a negative number to scroll backward, or 0 (zero) to not scroll at all. You must specify _Sheets_ if you don't specify _Position_.
    - `Position As Variant` (optional): Use xlFirst to scroll to the first sheet, or use xlLast to scroll to the last sheet. You must specify _Position_ if you don't specify _Sheets_.
- `SmallScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant]) As Variant`  
  Scrolls the contents of the window by rows or columns.
    - `Down As Variant` (optional): The number of rows to scroll the contents down.
    - `Up As Variant` (optional): The number of rows to scroll the contents up.
    - `ToRight As Variant` (optional): The number of columns to scroll the contents to the right.
    - `ToLeft As Variant` (optional): The number of columns to scroll the contents to the left.
- `PointsToScreenPixelsX(Points As Long) As Long`  
  Converts a horizontal measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a Long value.
    - `Points As Long` (required): The number of points horizontally along the top of the document window, starting from the left.
- `PointsToScreenPixelsY(Points As Long) As Long`  
  Converts a vertical measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a Long value.
    - `Points As Long` (required): The number of points vertically along the left edge of the document window, starting from the top.
- `RangeFromPoint(x As Long, y As Long) As Object`  
  Returns the Shape or Range object that is positioned at the specified pair of screen coordinates. If there isn't a shape located at the specified coordinates, this method returns Nothing.
    - `x As Long` (required): The value (in pixels) that represents the horizontal distance from the left edge of the screen, starting at the top.
    - `y As Long` (required): The value (in pixels) that represents the vertical distance from the top of the screen, starting on the left.
- `ScrollIntoView(Left As Long, Top As Long, Width As Long, Height As Long, [Start As Variant])`  
  Scrolls the document window so that the contents of a specified rectangular area are displayed in either the upper-left or lower-right corner of the document window or pane (depending on the value of the _Start_ argument).
    - `Left As Long` (required): The horizontal position of the rectangle (in points) from the left edge of the document window or pane.
    - `Top As Long` (required): The vertical position of the rectangle (in points) from the top of the document window or pane.
    - `Width As Long` (required): The width of the rectangle, in points.
    - `Height As Long` (required): The height of the rectangle, in points.
    - `Start As Variant` (optional): True to have the upper-left corner of the rectangle appear in the upper-left corner of the document window or pane. False to have the lower-right corner of the rectangle appear in the lower-right corner of the document window or pane. The default value is True.
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant], [PrToFileName As Variant]) As Variant`  
  Prints the object.
    - `From As Variant` (optional): The number of the page at which to start printing. If this argument is omitted, printing starts at the beginning.
    - `To As Variant` (optional): The number of the last page to print. If this argument is omitted, printing ends with the last page.
    - `Copies As Variant` (optional): The number of copies to print. If this argument is omitted, one copy is printed.
    - `Preview As Variant` (optional): True to have Microsoft Excel invoke print preview before printing the object. False (or omitted) to print the object immediately.
    - `ActivePrinter As Variant` (optional): Sets the name of the active printer.
    - `PrintToFile As Variant` (optional): True to print to a file. If _PrToFileName_ is not specified, Microsoft Excel prompts the user to enter the name of the output file.
    - `Collate As Variant` (optional): True to collate multiple copies.
    - `PrToFileName As Variant` (optional): If _PrintToFile_ is set to True, this argument specifies the name of the file that you want to print to.
