# Frameset

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209E2-0000-0000-C000-000000000046}  

Represents an entire frames page or a single frame on a frames page.

**Remarks:** Use the Frameset property of a Document or Pane object to return a Frameset object. - For properties or methods that affect all frames on a frames page, use the Frameset object from the Document object ( ActiveWindow.Document.Frameset). - For properties or methods that affect individual frames on a frames page, use the Frameset object from the Pane object ( ActiveWindow.ActivePane.Frameset). This example opens a file named "Proposal.doc," creates a frames page based on the file, and adds a frame (on the left side of the page) containing a table of contents for the file. This example adds a new frame to the right of the specified frame. This example sets the name of the third child Frameset object of the frames page to "BottomFrame." This example links the specified frame to a local file called "Order.htm." It sets the frame to be resizable, to appear with scrollbars in a web browser, and to be 25% as high as the active window. This example sets Microsoft Word to display frame borders in the specified frames page. This example sets the frame borders on the frames page to be 6 points wide and tan.

## Properties (20)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Frameset object.
- `_NewEnum As IUnknown  (read-only)`
- `ParentFrameset As Frameset  (read-only)`  
  Returns a Frameset object that represents the parent of the specified Frameset object on a frames page.
- `Type As WdFramesetType  (read-only)`  
  Returns the Frameset object type. Read-only WdFramesetType.
- `WidthType As WdFramesetSizeType  (read/write)`  
  Returns or sets the width type for the specified Frameset object. Read/write WdFramesetSizeType.
- `HeightType As WdFramesetSizeType  (read/write)`  
  Returns or sets the width type for the specified frame on a frames page. Read/write WdFramesetSizeType.
- `Width As Long  (read/write)`  
  Returns or sets the width (in points) of the specified Frameset object. Read/write Long.
- `Height As Long  (read/write)`  
  Returns or sets a Float that represents the height (in points) of the specified Frameset object. Read/write.
- `ChildFramesetCount As Long  (read-only)`  
  Returns the number of child Frameset objects associated with the specified Frameset object. Read-only Long.
- `ChildFramesetItem As Frameset  (read-only)`  
  Returns the Frameset object that represents the child Frameset object specified by the Index argument. Read-only.
- `FramesetBorderWidth As Single  (read/write)`  
  Returns or sets the width (in points) of the borders surrounding the frames on the specified frames page. Read/write Single.
- `FramesetBorderColor As WdColor  (read/write)`  
  Returns or sets the color of the frame borders on the specified frames page. Read/write.
- `FrameScrollbarType As WdScrollbarType  (read/write)`  
  Returns or sets when scroll bars are available for the specified frame when viewing its frames page in a web browser. Read/write WdScrollbarType.
- `FrameResizable As Boolean  (read/write)`  
  True if the user can resize the specified frame when the frames page is viewed in a web browser. Read/write Boolean.
- `FrameName As String  (read/write)`  
  Returns or sets the name of the specified frame on a frames page. Read/write String.
- `FrameDisplayBorders As Boolean  (read/write)`  
  True if the frame borders on the specified frames page are displayed. Read/write Boolean.
- `FrameDefaultURL As String  (read/write)`  
  Returns or sets the webpage or other document to be displayed in the specified frame when the frames page is opened. Read/write String.
- `FrameLinkToFile As Boolean  (read/write)`  
  True if the webpage or other document specified by the FrameDefaultURL property is an external file to which Microsoft Word maintains only a link from the specified frame. Read/write Boolean.

## Methods (2)

- `AddNewFrame(Where As WdFramesetNewFrameLocation) As Frameset`  
  Adds a new frame to a frames page.
    - `Where As WdFramesetNewFrameLocation` (required): Sets the location where the new frame is to be added in relation to the specified frame.
- `Delete()`  
  Deletes the specified Frameset object.
