# ColorEffect

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934E6-5A91-11CF-8700-00AA0060263B}  

Represents a color effect for an animation behavior.

**Example:**

```vba
Sub ChangeColorEffect()

    Dim sldFirst As Slide
    Dim shpHeart As Shape
    Dim effNew As Effect
    Dim bhvEffect As AnimationBehavior

    Set sldFirst = ActivePresentation.Slides(1)

    Set shpHeart = sldFirst.Shapes.AddShape(Type:=msoShapeHeart, _
        Left:=100, Top:=100, Width:=100, Height:=100)

    Set effNew = sldFirst.TimeLine.MainSequence.AddEffect _
        (Shape:=shpHeart, EffectID:=msoAnimEffectChangeFillColor, _
        Trigger:=msoAnimTriggerAfterPrevious)

    Set bhvEffect = effNew.Behaviors.Add(Type:=msoAnimTypeColor)

    With bhvEffect.ColorEffect
        .From.RGB = RGB(Red:=255, Green:=0, Blue:=0)
        .To.RGB = RGB(Red:=0, Green:=0, Blue:=255)
    End With

End Sub
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `By As ColorFormat  (read-only)`  
  Returns a ColorFormat object that represents a change to the color of the object by the specified number, expressed in RGB format. Read-only.
- `From As ColorFormat  (read-only)`  
  Sets or returns a ColorFormat object that represents the starting RGB color value of an animation behavior.
- `To As ColorFormat  (read-only)`  
  Sets or returns a ColorFormat object that represents the RGB color value of an animation behavior. Read/write.
