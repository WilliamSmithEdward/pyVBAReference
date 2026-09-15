# Timing

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934E0-5A91-11CF-8700-00AA0060263B}  

Represents timing properties for an animation effect.

**Remarks:** Use the following read/write properties of the Timing object to manipulate animation timing effects.

**Example:**

```vba
ActiveWindow.Selection.SlideRange(1).TimeLine.MainSequence(1).Timing.Duration = 5
```

## Properties (19)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Duration As Single  (read/write)`  
  Returns or sets the length of an animation in seconds. Read/write.
- `TriggerType As MsoAnimTriggerType  (read/write)`  
  Represents the trigger that starts an animation. Read/write.
- `TriggerDelayTime As Single  (read/write)`  
  Sets or returns the delay, in seconds, from when an animation trigger is enabled. Read/write.
- `TriggerShape As Shape  (read/write)`  
  Sets or returns a Shape object that represents the shape associated with an animation trigger. Read/write.
- `RepeatCount As Long  (read/write)`  
  Sets or returns the number of times to repeat an animation. Read/write.
- `RepeatDuration As Single  (read/write)`  
  Sets or returns how long repeated animations should last, in seconds. Read/write.
- `Speed As Single  (read/write)`  
  Returns or sets the speed, in seconds, of the specified animation. Read/write.
- `Accelerate As Single  (read/write)`  
  Returns or sets the percentage of the duration over which a timing acceleration should take place. Read/write.
- `Decelerate As Single  (read/write)`  
  Sets or returns the percentage of the duration over which a timing deceleration should take place. Read/write.
- `AutoReverse As MsoTriState  (read/write)`  
  Determines whether an effect should play forward and then in reverse, thereby doubling its duration. Read/write.
- `SmoothStart As MsoTriState  (read/write)`  
  Determines whether an animation should accelerate when it starts. Read/write.
- `SmoothEnd As MsoTriState  (read/write)`  
  Determines whether an animation should decelerate as it ends. Read/write.
- `RewindAtEnd As MsoTriState  (read/write)`  
  Represents whether an object returns to its beginning position after an animation has ended. Read/write.
- `Restart As MsoAnimEffectRestart  (read/write)`  
  Represents whether the animation effect restarts after the effect has started once. Read/write.
- `BounceEnd As MsoTriState  (read/write)`  
  Read/write
- `BounceEndIntensity As Single  (read/write)`  
  Read/write
- `TriggerBookmark As String  (read/write)`  
  Read/write
