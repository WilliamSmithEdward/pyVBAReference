# PlaySettings

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149348E-5A91-11CF-8700-00AA0060263B}  

Contains information about how the specified media clip will be played during a slide show.

**Example:**

```vba
Set clockMovie = ActivePresentation.Slides(1).Shapes _
    .AddMediaObject(FileName:="C:\WINNT\clock.avi", _
    Left:=20, Top:=20)
With clockMovie.AnimationSettings.PlaySettings
    .PlayOnEntry = True
    .PauseAnimation = False
    .HideWhileNotPlaying = True
End With
```

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `ActionVerb As String  (read/write)`  
  Returns or sets a string that contains the OLE verb that will be run when the specified OLE object is animated during a slide show. Read/write.
- `HideWhileNotPlaying As MsoTriState  (read/write)`  
  Determines whether the specified media clip is hidden during a slide show except when it is playing. Read/write.
- `LoopUntilStopped As MsoTriState  (read/write)`  
  Determines whether the specified movie or sound loops continuously until either the next movie or sound starts, the user clicks the slide, or a slide transition occurs. Read/write.
- `PlayOnEntry As MsoTriState  (read/write)`  
  Determines whether the specified movie or sound is played automatically when it is animated. Read/write.
- `RewindMovie As MsoTriState  (read/write)`  
  Determines whether the first frame of the specified movie is automatically redisplayed as soon as the movie has finished playing. Read/write.
- `PauseAnimation As MsoTriState  (read/write)`  
  Determines whether the slide show pauses until the specified media clip is finished playing. Read/write.
- `StopAfterSlides As Long  (read/write)`  
  Returns or sets the number of slides to be displayed before the media clip stops playing. Read/write.
