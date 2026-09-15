# SoundEffect

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493472-5A91-11CF-8700-00AA0060263B}  

Represents the sound effect that accompanies an animation or slide transition in a slide show.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes(1).AnimationSettings

    .TextLevelEffect = ppAnimateByAllLevels

    .SoundEffect.ImportFromFile "c:\sndsys\bass.wav"

End With
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write.
- `Type As PpSoundEffectType  (read/write)`  
  Represents the type of sound effect. Read/write.

## Methods (2)

- `ImportFromFile(FileName As String)`  
  Specifies the sound that will be played whenever the specified shape is clicked or animated or whenever the specified slide transition occurs.
- `Play()`  
  Plays the specified sound effect.
