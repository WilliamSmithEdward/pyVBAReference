# Zooms

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209A7-0000-0000-C000-000000000046}  

A collection of Zoom objects that represents the magnification options for each view (such as outline, normal, or print layout).

**Remarks:** Use the Zooms property to return the Zooms collection. The following example sets the zoom percentage for the active window to 100 percent in Normal view. The Add method isn't available for the Zooms collection. The Zooms collection includes a single Zoom object for each of the various view types (such as outline, normal, or page layout). You cannot enumerate the Zooms collection by using a For Each loop. Use Zooms (index), where index identifies the view type, to return a single Zoom object. The view type specified by index can be one of the following WdViewType constants: wdMasterView, wdNormalView, wdOutlineView, wdPrintPreview, wdPrintView, or wdWebView. The following example sets the magnification for the active window so that an entire page is visible. You can also use the Zoom property of the View object to return a single Zoom object. The following example sets the zoom percentage for the active window to 110 percent.

## Properties (3)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Zooms object.

## Methods (1)

- `Item(Index As WdViewType) As Zoom`  
  Returns the specified Zoom object.
    - `Index As WdViewType` (required): The specified zoom type.
