# EffectInformation

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934E2-5A91-11CF-8700-00AA0060263B}  

Represents various animation options for an Effect object.

**Remarks:** Use the members of the EffectInformation object to return the current state of an Effect object, such as the after effect, whether the background animates along with its corresponding text, whether text animates in reverse, play settings, sound effects, text building behavior. All of the members of the EffectInformation object are read-only. To change any effect information properties, you must use the methods of the corresponding Sequence object. Use the EffectInformationproperty of the Effect object to return an EffectInformation object.

**Example:**

```vba
Sub HideEffect()
    ActiveWindow.Selection.SlideRange(1).TimeLine _
        .MainSequence(1).EffectInformation.PlaySettings _
        .HideWhileNotPlaying = msoTrue
End Sub
```

## Properties (10)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `AfterEffect As MsoAnimAfterEffect  (read-only)`  
  Returns an PpAfterEffect constant that indicates whether an after effect appears dimmed, hidden, or unchanged after it runs. Read-only.
- `AnimateBackground As MsoTriState  (read-only)`  
  Returns msoTrue if the specified effect is a background animation. Read-only.
- `AnimateTextInReverse As MsoTriState  (read-only)`  
  Determines whether the specified shape is built in reverse order. Applies only to shapes (such as shapes containing lists) that can be built in more than one step. Read/write.
- `BuildByLevelEffect As MsoAnimateByLevel  (read-only)`  
  Determines the level of the animation build effect. Read-only.
- `Dim As ColorFormat  (read-only)`  
  Returns a ColorFormat object that represents the color to dim to after an animation is finished.
- `PlaySettings As PlaySettings  (read-only)`  
  Returns a PlaySettings object that contains information about how the specified media clip plays during a slide show. Read-only.
- `SoundEffect As SoundEffect  (read-only)`  
  Returns a SoundEffect object that represents the sound to be played during the transition to the specified slide. Read-only.
- `TextUnitEffect As MsoAnimTextUnitEffect  (read-only)`  
  Indicates whether the text in the specified shape is animated paragraph by paragraph, word by word, or letter by letter. Read-only.
