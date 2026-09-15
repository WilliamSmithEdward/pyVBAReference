# AnimationPoint

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934EB-5A91-11CF-8700-00AA0060263B}  

Represents an individual animation point for an animation behavior. The AnimationPoint object is a member of the AnimationPoints collection. The AnimationPoints collection contains all the animation points for an animation behavior.

**Example:**

```vba
Sub AniPoint()

    Dim sldNewSlide As Slide
    Dim shpHeart As Shape
    Dim effCustom As Effect
    Dim aniBehavior As AnimationBehavior
    Dim aptNewPoint As AnimationPoint

    Set sldNewSlide = ActivePresentation.Slides.Add _
        (Index:=1, Layout:=ppLayoutBlank)

    Set shpHeart = sldNewSlide.Shapes.AddShape _
        (Type:=msoShapeHeart, Left:=100, Top:=100, _
        Width:=200, Height:=200)

    Set effCustom = sldNewSlide.TimeLine.MainSequence _
        .AddEffect(shpHeart, msoAnimEffectCustom)

    Set aniBehavior = effCustom.Behaviors.Add(msoAnimTypeProperty)

    With aniBehavior.PropertyEffect
        .Property = msoAnimShapeFillColor
        Set aptNewPoint = .Points.Add
        aptNewPoint.Time = 0.2
        aptNewPoint.Value = RGB(0, 0, 0)
        Set aptNewPoint = .Points.Add
        aptNewPoint.Time = 0.5
        aptNewPoint.Value = RGB(0, 255, 0)
        Set aptNewPoint = .Points.Add
        aptNewPoint.Time = 1
        aptNewPoint.Value = RGB(0, 255, 255)
    End With

End Sub
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Time As Single  (read/write)`  
  Sets or returns the time at a given animation point. Read/write.
- `Value As Variant  (read/write)`  
  Sets or returns the value of a property for an animation point. Read/write.
- `Formula As String  (read/write)`  
  Returns or sets a String that represents a formula to use for calculating an animation. Read/write.

## Methods (1)

- `Delete()`  
  Deletes the specified AnimationPoint object.
