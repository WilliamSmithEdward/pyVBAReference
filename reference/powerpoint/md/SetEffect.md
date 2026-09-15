# SetEffect

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934F1-5A91-11CF-8700-00AA0060263B}  

Represents a set effect for an animation behavior. Use the SetEffect object to set the value of a property.

**Remarks:** Use the SetEffect property of the AnimationBehavior object to return a SetEffect object. Set effects can be changed using the SetEffect object's Property and To properties.

**Example:**

```vba
Sub ChangeSetEffect()

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

    Set bhvEffect = effNew.Behaviors.Add(msoAnimTypeSet)

    With bhvEffect.SetEffect
         .Property = msoAnimShapeFillColor
         .To = RGB(Red:=0, Green:=255, Blue:=255)
    End With

End Sub
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Property As MsoAnimProperty  (read/write)`  
  Sets or returns an MsoAnimProperty constant that represents an animation property. Read/write.
- `To As Variant  (read/write)`  
  Sets or returns a Variant that represents the value or ending value of the SetEffect object's Type property. Read/write.
