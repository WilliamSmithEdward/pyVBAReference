# AnimationBehavior

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934E4-5A91-11CF-8700-00AA0060263B}  

Represents the behavior of an animation effect, the main animation sequence, or an interactive animation sequence. The AnimationBehavior object is a member of the AnimationBehaviors collection.

**Example:**

```vba
Sub Change()
    With ActivePresentation.Slides(1).TimeLine.MainSequence(1) _
            .Behaviors(1).RotationEffect
        .From = 1
        .To = 180
    End With
End Sub
```

## Properties (14)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Additive As MsoAnimAdditive  (read/write)`  
  Sets or returns whether the current animation behavior is combined with other running animations. Read/write.
- `Accumulate As MsoAnimAccumulate  (read/write)`  
  Determines whether animation behaviors accumulate. Read/write.
- `Type As MsoAnimType  (read/write)`  
  Represents the type of animation. Read/write.
- `MotionEffect As MotionEffect  (read-only)`  
  Returns a MotionEffect object that represents the properties of a motion animation.
- `ColorEffect As ColorEffect  (read-only)`  
  Returns a ColorEffect object that represents the color properties for a specified animation behavior.
- `ScaleEffect As ScaleEffect  (read-only)`  
  Returns a ScaleEffect object for a given animation behavior. Read-only.
- `RotationEffect As RotationEffect  (read-only)`  
  Returns a RotationEffect object for an animation behavior. Read-only.
- `PropertyEffect As PropertyEffect  (read-only)`  
  Returns a PropertyEffect object for a given animation behavior. Read-only.
- `Timing As Timing  (read-only)`  
  Returns a Timing object that represents the timing properties for an animation sequence.
- `CommandEffect As CommandEffect  (read-only)`  
  Returns a CommandEffect object for the specified animation behavior. Read-only.
- `FilterEffect As FilterEffect  (read-only)`  
  Returns a FilterEffect object that represents a filter effect for an animation behavior. Read-only.
- `SetEffect As SetEffect  (read-only)`  
  Returns a SetEffect object for the animation behavior. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the specified AnimationBehavior object.
