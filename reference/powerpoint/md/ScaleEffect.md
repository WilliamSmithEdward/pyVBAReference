# ScaleEffect

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934E7-5A91-11CF-8700-00AA0060263B}  

Represents a scaling effect for an AnimationBehavior object.

**Example:**

```vba
ActivePresentation.Slides(1).TimeLine.MainSequence.Item.Behaviors(1).ScaleEffect
```

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `ByX As Single  (read/write)`  
  Sets or returns a Single that represents scaling or moving an object horizontally by a specified percentage of the screen width, depending on whether it used in conjunction with a ScaleEffect or MotionEffect object, respectively. For example, a value of 50 for a motion effect means to move the object half the screen width to the right. Read/write.
- `ByY As Single  (read/write)`  
  Sets or returns a Single that represents scaling or moving an object vertically by a specified percentage of the screen width, depending on whether it is used in conjunction with a ScaleEffect or MotionEffect object, respectively. Read/write.
- `FromX As Single  (read/write)`  
  Sets or returns a Single that represents the starting width or horizontal position of a ScaleEffect object, specified as a percent of the screen width. Read/write.
- `FromY As Single  (read/write)`  
  Returns or sets a Single that represents the starting height of a ScaleEffect object, specified as a percentage of the screen width. Read/write.
- `ToX As Single  (read/write)`  
  Sets or returns a Single that represents the ending width of a ScaleEffect object, specified as a percent of the screen width. Read/write.
- `ToY As Single  (read/write)`  
  Returns or sets a Single that represents the ending height of a ScaleEffect object, specified as a percentage of the screen width. Read/write.
