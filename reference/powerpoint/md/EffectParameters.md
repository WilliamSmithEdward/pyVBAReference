# EffectParameters

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934E1-5A91-11CF-8700-00AA0060263B}  

Represents various animation parameters for an Effect object, such as colors, fonts, sizes, and directions.

**Example:**

```vba
Sub effParam()

    Dim shpNew As Shape
    Dim effNew As Effect

    Set shpNew = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeHeart, Left:=100, _
        Top:=100, Width:=150, Height:=150)

    Set effNew = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpNew, EffectID:=msoAnimEffectChangeFillColor, _
        Trigger:=msoAnimTriggerAfterPrevious)

    With effNew.EffectParameters
        .Color1.RGB = RGB(Red:=0, Green:=0, Blue:=255)
        .Color2.RGB = RGB(Red:=255, Green:=0, Blue:=0)
    End With

End Sub
```

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Direction As MsoAnimDirection  (read/write)`  
  Determines the direction used for an animation effect. This property can be used only if the effect uses a direction. Read/write.
- `Amount As Single  (read/write)`  
  Returns or sets a Single that represents the number of degrees an animated shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation. Read/write.
- `Size As Single  (read/write)`  
  Returns or sets the character size, in points. Read/write.
- `Color2 As ColorFormat  (read-only)`  
  Returns a ColorFormat object that represents the color on which to end a color-cycle animation.
- `Relative As MsoTriState  (read/write)`  
  Determines whether to set the motion position relative to the position of the shape. Read/write.
- `FontName As String  (read/write)`  
  Returns or sets the name of the font in the specified WordArt. Read/write.
