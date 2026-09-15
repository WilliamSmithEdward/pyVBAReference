# SlideShowWindow

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493453-5A91-11CF-8700-00AA0060263B}  

Represents a window in which a slide show runs.

**Example:**

```vba
SlideShowWindows(2).Activate
```

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `View As SlideShowView  (read-only)`  
  Returns a SlideShowView object. Read-only.
- `Presentation As Presentation  (read-only)`  
  Returns a Presentation object that represents the presentation in which the specified document window or slide show window was created. Read-only.
- `IsFullScreen As MsoTriState  (read-only)`  
  Returns whether the specified slide show window occupies the entire screen. Read-only.
- `Left As Single  (read/write)`  
  Returns or sets a Single that represents the distance in points from the left edge of the document, application, and slide show windows to the left edge of the application window's client area. Setting this property to a very large positive or negative value may position the window completely off the desktop. Read/write.
- `Top As Single  (read/write)`  
  Returns or sets a Single that represents the distance in points from the top edge of the document, application, and slide show window to the top edge of the application window's client area. Read/write.
- `Width As Single  (read/write)`  
  Returns or sets the width of the specified object, in points. Read/write.
- `Height As Single  (read/write)`  
  Returns or sets the height of the specified object, in points. Read/write.
- `Active As MsoTriState  (read-only)`  
  Returns whether the specified pane or window is active. Read-only.
- `SlideNavigation As SlideNavigation  (read-only)`  
  Returns a SlideNavigation object that represents the slide navigation screen in slide show view. Read-only

## Methods (1)

- `Activate()`  
  Activates the specified object.
