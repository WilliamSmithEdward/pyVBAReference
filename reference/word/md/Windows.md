# Windows

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020961-0000-0000-C000-000000000046}  

A collection of Window objects that represent all the available windows. The Windows collection for the Application object contains all the windows in the application, whereas the Windows collection for the Document object contains only the windows that display the specified document.

**Remarks:** Use the Windows property to return the Windows collection. The following example tiles all the windows so that they don't overlap one another. Use the Add method or the NewWindow method to add a new window to the Windows collection. Each of the following statements creates a new window for the document in the active window. Use Windows (Index), where Index is the window name or the index number, to return a single Window object. The following example maximizes the Document1 window. The index number is the number to the left of the window name on the Window menu. The following example displays the caption of the first window in the Windows collection. A colon (:) and a number appear in the window caption when more than one window is open for a document. When you switch the view to print preview, a new window is created. This window is removed from the Windows collection when you close print preview.

## Properties (6)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of windows in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Windows object.
- `SyncScrollingSideBySide As Boolean  (read/write)`

## Methods (6)

- `Item(Index As Variant) As Window`  
  Returns an individual Window object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add([Window As Variant]) As Window`  
  Returns a Window object that represents a new window of a document.
    - `Window As Variant` (optional): The Window object you want to open another window for. If this argument is omitted, a new window is opened for the active document.
- `Arrange([ArrangeStyle As Variant])`  
  Arranges all open document windows in the application workspace.
    - `ArrangeStyle As Variant` (optional): The window arrangement. Can be either of the following WdArrangeStyle constants: wdIcons or wdTiled.
- `CompareSideBySideWith(Document As Variant) As Boolean`  
  Opens two windows in side by side mode. Returns a Boolean.
    - `Document As Variant` (required): The document to view in side by side windows.
- `BreakSideBySide() As Boolean`  
  Ends side by side mode if two windows are in side by side mode. Returns a Boolean that represents whether the method was successful.
- `ResetPositionsSideBySide()`  
  Resets two document windows that are in the Compare side by side with view mode.
