# PageSetup

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493466-5A91-11CF-8700-00AA0060263B}  

Contains information about the page setup for slides, notes pages, handouts, and outlines in a presentation.

**Example:**

```vba
With ActivePresentation.PageSetup

    .SlideWidth = 11 * 72

    .SlideHeight = 8.5 * 72

    .FirstSlideNumber = 17

End With
```

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `FirstSlideNumber As Long  (read/write)`  
  Returns or sets the slide number for the first slide in the presentation. Read/write.
- `SlideHeight As Single  (read/write)`  
  Returns or sets the slide height, in points. Read/write.
- `SlideWidth As Single  (read/write)`  
  Returns or sets the slide width, in points. Read/write.
- `SlideSize As PpSlideSizeType  (read/write)`  
  Returns or sets the slide size for the specified presentation. Read/write.
- `NotesOrientation As MsoOrientation  (read/write)`  
  Returns or sets the on-screen and printed orientation of notes pages, handouts, and outlines for the specified presentation. Read/write.
- `SlideOrientation As MsoOrientation  (read/write)`  
  Returns or sets the on-screen and printed orientation of slides in the specified presentation. Read/write.
