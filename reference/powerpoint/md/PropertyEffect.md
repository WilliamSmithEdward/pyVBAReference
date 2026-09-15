# PropertyEffect

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934E9-5A91-11CF-8700-00AA0060263B}  

Represents a property effect for an AnimationBehavior object.

**Example:**

```vba
ActivePresentation.Slides(1).TimeLine.MainSequence.Item(1) _
   .Behaviors(1).PropertyEffect
```

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Property As MsoAnimProperty  (read/write)`  
  Sets or returns an MsoAnimProperty constant that represents an animation property. Read/write.
- `Points As AnimationPoints  (read-only)`  
  Returns an AnimationPoints object that represents a point in an animation.
- `From As Variant  (read/write)`  
  Sets or returns a Variant that represents the starting value of an object's property. Read/write.
- `To As Variant  (read/write)`  
  Sets or returns a Variant that represents the ending value of an object's property. Read/write.
