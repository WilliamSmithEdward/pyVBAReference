# Effect

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934DF-5A91-11CF-8700-00AA0060263B}  

Represents timing information about a slide animation.

**Example:**

```vba
Sub NewShapeAndEffect()

    Dim shpStar As Shape

    Dim sldOne As Slide

    Dim effNew As Effect



    Set sldOne = ActivePresentation.Slides(1)

    Set shpStar = sldOne.Shapes.AddShape(Type:=msoShape5pointStar, _

        Left:=150, Top:=72, Width:=400, Height:=400)

    Set effNew = sldOne.TimeLine.MainSequence.AddEffect(Shape:=shpStar, _

        EffectId:=msoAnimEffectStretchy, Trigger:=msoAnimTriggerAfterPrevious)

    With effNew

        With .Behaviors.Add(msoAnimTypeScale).ScaleEffect

            .FromX = 75

            .FromY = 75

            .ToX = 0

            .ToY = 0

        End With

        .Timing.AutoReverse = msoTrue

    End With

End Sub
```

## Properties (14)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Shape As Shape  (read/write)`  
  Returns a Shape object that represents an animated shape.
- `Index As Long  (read-only)`  
  Returns a Long that represents the index number for an animation effect or design. Read-only.
- `Timing As Timing  (read-only)`  
  Returns a Timing object that represents the timing properties for an animation sequence.
- `EffectType As MsoAnimEffect  (read/write)`  
  Sets or returns an MsoAnimEffect constant that represents an animation effect type. Read/write.
- `EffectParameters As EffectParameters  (read-only)`  
  Returns an EffectParameters object that represents animation effect properties.
- `TextRangeStart As Long  (read-only)`  
  Returns or sets the start of a text range. Read-only.
- `TextRangeLength As Long  (read-only)`  
  Returns or sets a Long that represents the length of a text range. Read-only.
- `Paragraph As Long  (read/write)`  
  Returns or sets the paragraph in a text range to which to apply animation effects. Read/write.
- `DisplayName As String  (read-only)`  
  Returns the name of an animation effect. Read-only.
- `Exit As MsoTriState  (read/write)`  
  Determines whether the animation effect is an exit effect. Read/write.
- `Behaviors As AnimationBehaviors  (read-only)`  
  Returns a specified slide animation behavior as an AnimationBehaviors collection.
- `EffectInformation As EffectInformation  (read-only)`  
  Returns an EffectInformation object that represents information for a specified animation effect.

## Methods (4)

- `MoveTo(toPos As Long)`  
  Moves the specified object to a specific location within the same collection, renumbering all other items in the collection appropriately.
    - `toPos As Long` (required): The index position to which to move the animation effect.
- `MoveBefore(Effect As Effect)`  
  Moves one animation effect to before another animation effect.
- `MoveAfter(Effect As Effect)`  
  Moves one animation effect to after another animation effect.
- `Delete()`  
  Deletes the specified Effect object.
