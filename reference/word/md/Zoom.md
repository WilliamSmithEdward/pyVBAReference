# Zoom

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209A6-0000-0000-C000-000000000046}  

Contains magnification options (for example, the zoom percentage) for a window or pane. The Zoom object is a member of the Zooms collection.

**Remarks:** Use the Zoom property of the View object to return a single Zoom object. The following example sets the zoom percentage for the active window to 110 percent. Use Zooms (Index), where Index identifies the view type, to return a single Zoom object. The view type specified by index can be one of the following WdViewType constants: wdMasterView, wdNormalView, wdOutlineView, wdPrintPreview, wdPrintView, or wdWebView. The following example sets the magnification for the active window so that an entire page is visible. The Add method isn't available for the Zooms collection. The Zooms collection includes a single Zoom object for each of the various view types (such as outline, normal, or page layout).

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Zoom object.
- `Percentage As Long  (read/write)`  
  Returns or sets the magnification for a window as a percentage. Read/write Long.
- `PageFit As WdPageFit  (read/write)`  
  Returns or sets the view magnification of a window so that either the entire page is visible or the entire width of the page is visible. Read/write WdPageFit.
- `PageRows As Long  (read/write)`  
  Returns or sets the number of pages to be displayed one above the other on-screen at the same time in print layout view or print preview. Read/write Long.
- `PageColumns As Long  (read/write)`  
  Returns or sets the number of pages to be displayed side by side on-screen at the same time in print layout view or print preview. Read/write Long.
