# SlideShowView

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493459-5A91-11CF-8700-00AA0060263B}  

Represents the view in a slide show window.

**Example:**

```vba
SlideShowWindows(1).View.First
```

## Properties (21)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Zoom As Long  (read-only)`  
  Returns the zoom setting of the specified slide show window view as a percentage of normal size. Read-only.
- `Slide As Slide  (read-only)`  
  Returns a Slide object that represents the slide that's currently displayed in the specified slide show window view. Read-only.
- `PointerType As PpSlideShowPointerType  (read/write)`  
  Returns or sets the type of pointer used in the slide show. Read/write.
- `State As PpSlideShowState  (read/write)`  
  Returns or sets the state of the slide show. Read/write.
- `AcceleratorsEnabled As MsoTriState  (read/write)`  
  Determines whether shortcut keys are enabled during a slide show. Read/write.
- `PresentationElapsedTime As Single  (read-only)`  
  Returns the number of seconds that have elapsed since the beginning of the specified slide show. Read-only.
- `SlideElapsedTime As Single  (read/write)`  
  Returns the number of seconds that the current slide has been displayed. Read/write.
- `LastSlideViewed As Slide  (read-only)`  
  Returns a Slide object that represents the slide viewed immediately before the current slide in the specified slide show view.
- `AdvanceMode As PpSlideShowAdvanceMode  (read-only)`  
  Returns a value that indicates how the slide show in the specified view advances. Read-only.
- `PointerColor As ColorFormat  (read-only)`  
  Returns a ColorFormat object that represents the pointer color for the specified presentation during one slide show. Read-only.
- `IsNamedShow As MsoTriState  (read-only)`  
  Determines whether a custom (named) slide show is displayed in the specified slide show view. Read-only.
- `SlideShowName As String  (read-only)`  
  Returns the name of the custom slide show that's currently running in the specified slide show view. Read-only.
- `CurrentShowPosition As Long  (read-only)`  
  Returns the position of the current slide within the slide show that is showing in the specified view. Read-only.
- `MediaControlsVisible As MsoTriState  (read-only)`  
  Indicates whether the media controls are visible. Read-only.
- `MediaControlsLeft As Single  (read-only)`  
  Returns the distance, in points, from the left edge of the media control bounding box to the left edge of the Slide. Read-only.
- `MediaControlsTop As Single  (read-only)`  
  Returns the distance, in points, from the top edge of the media control bounding box to the top edge of the Slide object. Read-only.
- `MediaControlsWidth As Single  (read-only)`  
  Returns the width, in points, of the media control bounding box. Read-only.
- `MediaControlsHeight As Single  (read-only)`  
  Returns the height of the media control bounding box. Read-only.
- `LaserPointerEnabled As Boolean  (read/write)`  
  Returns True if the current slide show pointer is a laser pointer. This property is applicable only while the slide show is running. Read/write.

## Methods (16)

- `DrawLine(BeginX As Single, BeginY As Single, EndX As Single, EndY As Single)`  
  Draws a line in the specified slide show view.
    - `BeginX As Single` (required): The position (in points) of the line's starting point relative to the upper-left corner of the slide.
    - `BeginY As Single` (required): The position (in points) of the line's starting point relative to the upper-left corner of the slide.
    - `EndX As Single` (required): The position (in points) of the line's ending point relative to the upper-left corner of the slide.
    - `EndY As Single` (required): The position (in points) of the line's ending point relative to the upper-left corner of the slide.
- `EraseDrawing()`  
  Removes lines drawn during a slide show by using either the DrawLine method or the pen tool.
- `First()`  
  Sets the specified slide show view to display the first slide in the presentation.
- `Last()`  
  Sets the specified slide show view to display the last slide in the presentation.
- `Next()`  
  Displays the slide immediately following the slide that's currently displayed.
- `Previous()`  
  Shows the slide immediately preceding the slide that's currently displayed.
- `GotoSlide(Index As Long, [ResetSlide As MsoTriState])`  
  Switches to the specified slide during a slide show. You can specify whether you want the animation effects to be rerun.
    - `Index As Long` (required): The number of the slide to switch to.
    - `ResetSlide As MsoTriState` (optional): Whether animation effects should be rerun when returning to the first slide. See Remarks for more information.
- `GotoNamedShow(SlideShowName As String)`  
  Switches to the specified custom, or named, slide show during another slide show. When the slide show advances from the current slide, the next slide displayed will be the next one in the specified custom slide show, not the next one in current slide show.
    - `SlideShowName As String` (required): The name of the custom slide show to be switched to.
- `EndNamedShow()`  
  Switches from running a custom, or named, slide show to running the entire presentation of which the custom show is a subset. When the slide show advances from the current slide, the next slide displayed will be the next one in the entire presentation, not the next one in the custom slide show.
- `ResetSlideTime()`  
  Resets the elapsed time (represented by the SlideElapsedTime property) for the slide that's currently displayed to 0 (zero).
- `Exit()`  
  Ends the specified slide show.
- `GotoClick(Index As Long)`  
  Plays an animation associated with a specified mouse click and any animations that follow on the slide.
    - `Index As Long` (required): The index number of the mouse click that initiates an animation.
- `GetClickIndex() As Long`  
  Returns the index number of the current mouse click for an animation that is actively playing on a slide or has just finished.
- `GetClickCount() As Long`  
  Returns the number of mouse clicks that are defined for a slide.
- `FirstAnimationIsAutomatic() As Boolean`  
  Returns True if the current slide has an initial animation that runs automatically.
- `Player(ShapeId As Variant) As Player`  
  Allows access to playback controls for the associated view in the current window.
    - `ShapeId As Variant` (required): The playback control.
