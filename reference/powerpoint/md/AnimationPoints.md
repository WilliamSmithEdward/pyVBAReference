# AnimationPoints

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934EA-5A91-11CF-8700-00AA0060263B}  

Represents a collection of animation points for a PropertyEffect object.

**Example:**

```vba
Sub AddPoint()
    ActivePresentation.Slides(1).TimeLine.MainSequence(1) _
        .Behaviors(1).PropertyEffect.Points.Add
End Sub
```

## Properties (4)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Smooth As MsoTriState  (read/write)`  
  Determines whether the transition from one animation point to another is smoothed. Read/write.

## Methods (2)

- `Item(Index As Long) As AnimationPoint`  
  Returns a single AnimationPoint object from the specified AnimationPoints collection.
    - `Index As Long` (required): The index number of the single AnimationPoint object in the collection to be returned.
- `Add([Index As Long]) As AnimationPoint`  
  Returns an AnimationPoint object that represents a new animation point.
    - `Index As Long` (optional): The position of the animation point in relation to other animation points. The default value is -1, which means that if you omit the Index parameter, the new animation point is added to the end of existing animation points.
