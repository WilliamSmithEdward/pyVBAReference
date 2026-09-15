# FilterEffect

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934F0-5A91-11CF-8700-00AA0060263B}  

Represents a filter effect for an animation behavior.

**Remarks:** Use the FilterEffect property of the AnimationBehavior object to return a FilterEffect object. Filter effects can be changed using the FilterEffect object's Reveal, SubType, and Type properties.

**Example:**

```vba
Sub ChangeFilterEffect()

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

    Set bhvEffect = effNew.Behaviors.Add(msoAnimTypeFilter)

    With bhvEffect.FilterEffect
         .Type = msoAnimFilterEffectTypeWipe
         .Subtype = msoAnimFilterEffectSubtypeUp
         .Reveal = msoTrue
    End With

End Sub
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Type As MsoAnimFilterEffectType  (read/write)`  
  Represents the type of animation. Read/write.
- `Subtype As MsoAnimFilterEffectSubtype  (read/write)`  
  Sets or returns the subtype of the filter effect. Read/write.
- `Reveal As MsoTriState  (read/write)`  
  Determines how the embedded objects will be revealed. Read/write.
