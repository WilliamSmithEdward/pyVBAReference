# SlideShowSettings

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149345A-5A91-11CF-8700-00AA0060263B}  

Represents the slide show setup for a presentation.

**Example:**

```vba
For Each s In ActivePresentation.Slides

    With s.SlideShowTransition

        .AdvanceOnTime = True

        .AdvanceTime = 5

    End With

Next



With ActivePresentation.SlideShowSettings

    .RangeType = ppShowSlideRange

    .StartingSlide = 2

    .EndingSlide = 4

    .AdvanceMode = ppSlideShowUseSlideTimings

    .LoopUntilStopped = True

    .Run

End With
```

## Properties (16)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `PointerColor As ColorFormat  (read-only)`  
  Returns the pointer color for the specified presentation as a ColorFormat object. Read-only.
- `NamedSlideShows As NamedSlideShows  (read-only)`  
  Returns a NamedSlideShows collection that represents all the named slide shows (custom slide shows) in the specified presentation. Read-only.
- `StartingSlide As Long  (read/write)`  
  Returns or sets the first slide to be displayed in the specified slide show. Read/write.
- `EndingSlide As Long  (read/write)`  
  Returns or sets the last slide to be displayed in the specified slide show. Read/write.
- `AdvanceMode As PpSlideShowAdvanceMode  (read/write)`  
  Returns or sets a value that indicates how the slide show advances. Read/write.
- `LoopUntilStopped As MsoTriState  (read/write)`  
  Determines whether specified slide show loops continuously until the user presses ESC. Read/write.
- `ShowType As PpSlideShowType  (read/write)`  
  Returns or sets the show type for the specified slide show. Read/write.
- `ShowWithNarration As MsoTriState  (read/write)`  
  Determines whether the specified slide show is shown with narration. Read/write.
- `ShowWithAnimation As MsoTriState  (read/write)`  
  Determines whether the specified slide show displays shapes with assigned animation settings. Read/write.
- `SlideShowName As String  (read/write)`  
  Returns or sets the name of the custom slide show to run in response to a mouse action on the shape during a slide show. Read/write.
- `RangeType As PpSlideShowRangeType  (read/write)`  
  Returns or sets the type of slide show to run. Read/write.
- `ShowScrollbar As MsoTriState  (read/write)`  
  Determines whether to display the scroll bar during a slide show in browse mode. Read/write.
- `ShowPresenterView As MsoTriState  (read/write)`  
  Returns the presenter view of the SlideShowSettings object. Read/write.
- `ShowMediaControls As MsoTriState  (read/write)`  
  Allows access to the media controls in the SlideShowSettings object. Read/write.

## Methods (1)

- `Run() As SlideShowWindow`  
  Runs a slide show of the specified presentation. Returns a SlideShowWindow object.
