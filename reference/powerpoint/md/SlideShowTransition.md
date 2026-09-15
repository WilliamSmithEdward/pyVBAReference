# SlideShowTransition

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493471-5A91-11CF-8700-00AA0060263B}  

Contains information about how the specified slide advances during a slide show.

**Example:**

```vba
With ActivePresentation.Slides(1).SlideShowTransition

    .Speed = ppTransitionSpeedFast

    .EntryEffect = ppEffectStripsDownLeft

    .SoundEffect.ImportFromFile "c:\sndsys\bass.wav"

    .AdvanceOnTime = True

    .AdvanceTime = 5

End With

ActivePresentation.SlideShowSettings.AdvanceMode = _

    ppSlideShowUseSlideTimings
```

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `AdvanceOnClick As MsoTriState  (read/write)`  
  Determines whether the specified slide advances when it is clicked during a slide show. Read/write.
- `AdvanceOnTime As MsoTriState  (read/write)`  
  Determines whether the specified slide advances automatically after a specified amount of time has elapsed. Read/write.
- `AdvanceTime As Single  (read/write)`  
  Returns or sets the amount of time, in seconds, after which the specified slide transition will occur. Read/write.
- `EntryEffect As PpEntryEffect  (read/write)`  
  Returns or sets the special effect applied to the specified slide transition. Read/write.
- `Hidden As MsoTriState  (read/write)`  
  Determines whether the specified slide is hidden during a slide show. Read/write.
- `LoopSoundUntilNext As MsoTriState  (read/write)`  
  Specifies whether the sound that's been set for the specified slide transition loops until the next sound starts. Read/write.
- `SoundEffect As SoundEffect  (read-only)`  
  Returns a SoundEffect object that represents the sound to be played during the transition to the specified slide. Read-only.
- `Speed As PpTransitionSpeed  (read/write)`  
  Represents the speed of the transition to the specified slide. Read/write.
- `Duration As Single  (read/write)`  
  Returns or sets the length of an animation in seconds. Read/write.
