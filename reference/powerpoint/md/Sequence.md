# Sequence

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934DE-5A91-11CF-8700-00AA0060263B}  

Represents a collection of Effect objects for a slide's interactive animation sequences. The Sequence collection is a member of the Sequences collection.

**Example:**

```vba
Sub NewEffect()

    Dim effNew As Effect
    Dim shpFirst As Shape

    Set shpFirst = ActivePresentation.Slides(1).Shapes(1)

    Set effNew = ActivePresentation.Slides(1).TimeLine.MainSequence.AddEffect _
        (Shape:=shpFirst, effectId:=msoAnimEffectBlinds)

End Sub
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (11)

- `Item(Index As Long) As Effect`  
  Returns a single Effect object from the specified Sequence collection.
    - `Index As Long` (required): The index number of the single Effect object in the collection to be returned.
- `AddEffect(Shape As Shape, effectId As MsoAnimEffect, [Level As MsoAnimateByLevel], [trigger As MsoAnimTriggerType], [Index As Long]) As Effect`  
  Returns an Effect object that represents a new animation effect added to a sequence of animation effects.
    - `Shape As Shape` (required): The shape to which the animation effect is added.
    - `effectId As MsoAnimEffect` (required): The animation effect to be applied.
    - `Level As MsoAnimateByLevel` (optional): For charts, diagrams, or text, the level to which the animation effect will be applied. The default value is msoAnimationLevelNone.
    - `trigger As MsoAnimTriggerType` (optional): The action that triggers the animation effect. The default value is msoAnimTriggerOnPageClick.
    - `Index As Long` (optional): The position at which the effect will be placed in the collection of animation effects. The default value is -1 (added to the end).
- `Clone(Effect As Effect, [Index As Long]) As Effect`  
  Creates a copy of an Effect object, and adds it to the Sequences collection at the specified index position.
    - `Effect As Effect` (required): Effect object. The animation effect to be cloned.
    - `Index As Long` (optional): The position at which the cloned animation effect will be added to the Sequences collection. The default value is -1 (added to the end).
- `FindFirstAnimationFor(Shape As Shape) As Effect`  
  Returns an Effect object that represents the first animation for a given shape.
    - `Shape As Shape` (required): The shape for which to find the first animation.
- `FindFirstAnimationForClick(click As Long) As Effect`  
  Returns an Effect object that represents the first animation started by the specified click number.
    - `click As Long` (required): The specified click number.
- `ConvertToBuildLevel(Effect As Effect, Level As MsoAnimateByLevel) As Effect`  
  Changes the build level information for a specified animation effect. Returns an Effect object that represents the build level information.
    - `Effect As Effect` (required): The specified animation effect.
    - `Level As MsoAnimateByLevel` (required): The animation build level.
- `ConvertToAfterEffect(Effect As Effect, After As MsoAnimAfterEffect, [DimColor As MsoRGBType], [DimSchemeColor As PpColorSchemeIndex]) As Effect`  
  Specifies what an effect should do after it is finished. Returns an Effect object that represents an after effect.
    - `Effect As Effect` (required): The effect to which the after effect will be added.
    - `After As MsoAnimAfterEffect` (required): The behavior of the after effect.
    - `DimColor As MsoRGBType` (optional): A single color to apply the after effect.
    - `DimSchemeColor As PpColorSchemeIndex` (optional): A predefined color scheme to apply to the after effect.
- `ConvertToAnimateBackground(Effect As Effect, AnimateBackground As MsoTriState) As Effect`  
  Determines whether the background will be animated separately from, or in addition to, its accompanying text. Returns an Effect object representing the newly-modified animation effect.
    - `Effect As Effect` (required): The animation effect to be applied to the background.
    - `AnimateBackground As MsoTriState` (required): Determines whether the text will be animated separately from the background.
- `ConvertToAnimateInReverse(Effect As Effect, animateInReverse As MsoTriState) As Effect`  
  Determines whether text will be animated in reverse order. Returns an Effect object representing the text animation.
    - `Effect As Effect` (required): The animation effect to which the reversal will apply.
    - `animateInReverse As MsoTriState` (required): Determines the text animation order.
- `ConvertToTextUnitEffect(Effect As Effect, unitEffect As MsoAnimTextUnitEffect) As Effect`  
  Returns an Effect object that represents how text should be animated.
    - `Effect As Effect` (required): The animation effect to which the text unit effect applies.
    - `unitEffect As MsoAnimTextUnitEffect` (required): How the text should be animated.
- `AddTriggerEffect(pShape As Shape, effectId As MsoAnimEffect, trigger As MsoAnimTriggerType, pTriggerShape As Shape, [bookmark As String], [Level As MsoAnimateByLevel]) As Effect`  
  Adds a trigger effect to the animation in a Sequence object.
    - `pShape As Shape` (required): The Shape object with animation.
    - `effectId As MsoAnimEffect` (required): The type of animation.
    - `trigger As MsoAnimTriggerType` (required): The type of trigger effect to add.
    - `pTriggerShape As Shape` (required): The Shape object that represents the trigger.
    - `bookmark As String` (optional): The bookmark.
    - `Level As MsoAnimateByLevel` (optional): The level of animation.
