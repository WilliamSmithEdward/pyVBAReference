# RotationEffect

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934E8-5A91-11CF-8700-00AA0060263B}  

Represents a rotation effect for an AnimationBehavior object.

**Example:**

```vba
ActivePresentation.Slides(1).TimeLine.MainSequence.Item.Behaviors(1).RotationEffect
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `By As Single  (read/write)`  
  Sets or returns a Single that represents the rotation of an object by the specified number of degrees; for example, a value of 180 means to rotate the object by 180 degrees. Read/write.
- `From As Single  (read/write)`  
  Sets or returns a Single that represents the starting angle in degrees, specified relative to the screen (for example, 90 degrees is completely horizontal). Read/write.
- `To As Single  (read/write)`  
  Sets or returns a Single that represents the ending rotation of an object in degrees, specified relative to the screen (for example, 90 degrees is completely horizontal). Read/write.
