# Panes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002095F-0000-0000-C000-000000000046}  

A collection of Pane objects that represent the window panes for a single window.

**Remarks:** Use the Panes property to return the Panes collection. The following example splits the active window and hides the ruler for each pane. Use the Add method or the Split property to add a window pane. The following example splits the active window at 20 percent of the current window size. The following example splits the active window in half. Use the SplitSpecial property to show comments, footnotes, or endnotes in a separate pane. A window has more than one pane if it is split, or if the active view isn't print layout view and information such as footnotes or comments is displayed. The following example displays the footnote pane in normal view and then prompts the user to close the pane.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of panes in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Panes object.

## Methods (2)

- `Item(Index As Long) As Pane`  
  Returns the specified pane as a Pane object.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add([SplitVertical As Variant]) As Pane`  
  Returns a Pane object that represents a new pane to a window.
    - `SplitVertical As Variant` (optional): A number that represents the percentage of the window, from top to bottom, you want to appear above the split.
