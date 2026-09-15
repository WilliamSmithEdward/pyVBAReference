# DocumentWindow

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493457-5A91-11CF-8700-00AA0060263B}  

Represents a document window. The DocumentWindow object is a member of the DocumentWindows collection. The DocumentWindows collection contains all the open document windows.

**Remarks:** Use the Presentation property to return the presentation that's currently running in the specified document window. Use the Selection property to return the selection. Use the SplitHorizontal property to return the percentage of the screen width that the outline pane occupies in normal view. Use the SplitVertical property to return the percentage of the screen height that the slide pane occupies in normal view. Use the View property to return the view in the specified document window.

**Example:**

```vba
Windows(2).Activate
```

## Properties (18)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Selection As Selection  (read-only)`  
  Returns a Selection object that represents the selection in the specified document window. Read-only.
- `View As View  (read-only)`  
  Returns a View object that represents the view in the specified document window. Read-only.
- `Presentation As Presentation  (read-only)`  
  Returns a Presentation object that represents the presentation in which the specified document window or slide show window was created. Read-only.
- `ViewType As PpViewType  (read/write)`  
  Returns or sets the type of the view contained in the specified document window. Read/write.
- `BlackAndWhite As MsoTriState  (read/write)`  
  Determines whether the document window display is black and white. Read/write.
- `Active As MsoTriState  (read-only)`  
  Returns whether the specified pane or window is active. Read-only.
- `WindowState As PpWindowState  (read/write)`  
  Returns or sets the state of the specified window. Read/write.
- `Caption As String  (read-only)`  
  Returns the text that appears in the title bar of the document window. Read-only.
- `Left As Single  (read/write)`  
  Returns or sets a Single that represents the distance in points from the left edge of the document, application, and slide show windows to the left edge of the application window's client area. Setting this property to a very large positive or negative value may position the window completely off the desktop. Read/write.
- `Top As Single  (read/write)`  
  Returns or sets a Single that represents the distance in points from the top edge of the document, application, and slide show window to the top edge of the application window's client area. Read/write.
- `Width As Single  (read/write)`  
  Returns or sets the width of the specified object, in points. Read/write.
- `Height As Single  (read/write)`  
  Returns or sets the height of the specified object, in points. Read/write.
- `ActivePane As Pane  (read-only)`  
  Returns a Pane object that represents the active pane in the document window. Read-only.
- `Panes As Panes  (read-only)`  
  Returns a Panes collection that represents the panes in the document window. Read-only.
- `SplitVertical As Long  (read/write)`  
  Returns or sets the percentage of the document window height that the slide pane occupies in normal view. Corresponds to the pane divider position between the slide and notes panes. Read/write.
- `SplitHorizontal As Long  (read/write)`  
  Returns or sets the percentage of the document window width that the outline pane occupies in normal view. Corresponds to the pane divider position between the slide and outline panes. Read/write.

## Methods (13)

- `FitToPage()`  
  Adjusts the size of the specified document window to accommodate the information that's currently displayed.
- `Activate()`  
  Activates the specified object.
- `LargeScroll([Down As Long], [Up As Long], [ToRight As Long], [ToLeft As Long])`  
  Scrolls through the specified document window by pages.
    - `Down As Long` (optional): Specifies the number of pages to scroll down.
    - `Up As Long` (optional): Specifies the number of pages to scroll up.
    - `ToRight As Long` (optional): Specifies the number of pages to scroll right.
    - `ToLeft As Long` (optional): Specifies the number of pages to scroll left.
- `SmallScroll([Down As Long], [Up As Long], [ToRight As Long], [ToLeft As Long])`  
  Scrolls through the specified document window by lines and columns.
    - `Down As Long` (optional): Specifies the number of lines to scroll down.
    - `Up As Long` (optional): Specifies the number of lines to scroll up.
    - `ToRight As Long` (optional): Specifies the number of columns to scroll right.
    - `ToLeft As Long` (optional): Specifies the number of columns to scroll left.
- `NewWindow() As DocumentWindow`  
  Opens a new window that contains the same document that is displayed in the specified window. Returns a DocumentWindow object that represents the new window.
- `Close()`  
  Closes the specified document window.
- `RangeFromPoint(X As Long, Y As Long) As Object`  
  Returns the Shape object that is located at the point specified by the screen position coordinate pair. If no shape is located at the coordinate pair specified, then the method returns Nothing.
    - `X As Long` (required): The horizontal distance (in pixels) from the left edge of the screen to the point.
    - `Y As Long` (required): The vertical distance (in pixels) from the top of the screen to the point.
- `PointsToScreenPixelsX(Points As Single) As Long`  
  Converts a horizontal measurement from points to pixels. Used to return a horizontal screen location for a text frame or shape. Returns the converted measurement as a Single.
    - `Points As Single` (required): The horizontal measurement (in points) to be converted to pixels.
- `PointsToScreenPixelsY(Points As Single) As Long`  
  Converts a vertical measurement from points to pixels. Used to return a vertical screen location for a text frame or shape. Returns the converted measurement as a Single.
    - `Points As Single` (required): The vertical measurement (in points) to be converted to pixels.
- `ScrollIntoView(Left As Single, Top As Single, Width As Single, Height As Single, [Start As MsoTriState])`  
  Scrolls the document window so that items within a specified rectangular area are displayed in the document window or pane.
    - `Left As Single` (required): The horizontal distance (in points) from the left edge of the document window to the rectangle.
    - `Top As Single` (required): The vertical distance (in points) from the upper part of the document window to the rectangle.
    - `Width As Single` (required): The width of the rectangle (in points).
    - `Height As Single` (required): The height of the rectangle (in points).
    - `Start As MsoTriState` (optional): Determines the starting position of the rectangle in relation to the document window.
- `IsSectionExpanded(sectionIndex As Long) As Boolean`  
  Indicates whether the selected section is expanded in the DocumentWindow.
    - `sectionIndex As Long` (required): The index of the section.
- `ExpandSection(sectionIndex As Long, Expand As Boolean)`  
  Expands the section in the current DocumentWindow.
    - `sectionIndex As Long` (required): The index of the section to be expanded.
    - `Expand As Boolean` (required): Indicates whether the section was expanded.
- `ShowInsertAppDialog(Filter As Long)`
