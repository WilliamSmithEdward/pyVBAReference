# AnimationSettings

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149348B-5A91-11CF-8700-00AA0060263B}  

Represents the special effects applied to the animation for the specified shape during a slide show.

**Example:**

```vba
Set sObjs = ActivePresentation.Slides.Add(2, ppLayoutText).Shapes

sObjs.Title.TextFrame.TextRange.Text = "Top Three Reasons"

With sObjs.Placeholders(2)

    .TextFrame.TextRange.Text = _

        "Reason 1" & VBNewLine & "Reason 2" & VBNewLine & "Reason 3"

    With .AnimationSettings

        .TextLevelEffect = ppAnimateByFirstLevel

        .EntryEffect = ppEffectFlyFromLeft

        .AfterEffect = ppAfterEffectDim

        .DimColor.RGB = RGB(100, 120, 100)

        .AnimateTextInReverse = True

    End With

End With
```

## Properties (16)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `DimColor As ColorFormat  (read-only)`  
  Returns or sets a ColorFormat object that represents the color of the specified shape after it is been built. Read-only.
- `SoundEffect As SoundEffect  (read-only)`  
  Returns a SoundEffect object that represents the sound to be played during the transition to the specified slide. REad-only.
- `EntryEffect As PpEntryEffect  (read/write)`  
  For the AnimationSettings object, this property returns or sets the special effect applied to the animation for the specified shape. Read/write.
- `AfterEffect As PpAfterEffect  (read/write)`  
  Returns or sets a PpAfterEffect constant that indicates whether the specified shape appears dimmed, hidden, or unchanged after it is built. Read/write.
- `AnimationOrder As Long  (read/write)`  
  Returns or sets an integer that represents the position of the specified shape within the collection of shapes to be animated. Read/write.
- `AdvanceMode As PpAdvanceMode  (read/write)`  
  Returns or sets a value that indicates whether the specified shape animation advances only when clicked or automatically after a specified amount of time. Read/write.
- `AdvanceTime As Single  (read/write)`  
  Returns or sets the amount of time, in seconds, after which the specified shape will become animated. Read/write.
- `PlaySettings As PlaySettings  (read-only)`  
  Returns a PlaySettings object that contains information about how the specified media clip plays during a slide show. Read-only.
- `TextLevelEffect As PpTextLevelEffect  (read/write)`  
  Indicates whether the text in the specified shape is animated by first-level paragraphs, second-level paragraphs, or some other level of paragraphs (up to fifth-level paragraphs). Read/write.
- `TextUnitEffect As PpTextUnitEffect  (read/write)`  
  Indicates whether the text in the specified shape is animated paragraph by paragraph, word by word, or letter by letter. Read/write.
- `Animate As MsoTriState  (read/write)`  
  Determines whether the specified shape is animated during a slide show. Read/write.
- `AnimateBackground As MsoTriState  (read/write)`  
  If the specified object is an AutoShape, specifies if the shape is animated separately from the text it contains. Read/write.
- `AnimateTextInReverse As MsoTriState  (read/write)`  
  Determines whether the specified shape is built in reverse order. Applies only to shapes (such as shapes containing lists) that can be built in more than one step. Read/write.
- `ChartUnitEffect As PpChartUnitEffect  (read/write)`  
  Returns or sets a value that indicates whether the graph range is animated by series, category, or element. Read/write.
