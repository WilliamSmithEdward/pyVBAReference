# AnimationBehaviors

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934E3-5A91-11CF-8700-00AA0060263B}  

Represents a collection of AnimationBehavior objects.

**Example:**

```vba
Sub AnimationObject()

    Dim timeMain As TimeLine



    'Reference the main animation timeline

    Set timeMain = ActivePresentation.Slides(1).TimeLine



    'Add a five-second animated rotation behavior

    'as the first animation in the main animation sequence

    timeMain.MainSequence(1).Behaviors.Add Type:=msoAnimTypeRotation, Index:=1

End Sub
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (2)

- `Item(Index As Long) As AnimationBehavior`  
  Returns a single AnimationBehavior object from the specified AnimationBehaviors collection.
    - `Index As Long` (required): The index number of the single AnimationBehavior object in the collection to be returned.
- `Add(Type As MsoAnimType, [Index As Long]) As AnimationBehavior`  
  Returns an AnimationBehavior object that represents a new animation behavior.
    - `Type As MsoAnimType` (required): The type of the animation behavior.
    - `Index As Long` (optional): The position of the animation behavior in relation to other animation behaviors. The default value is -1, which means that if you omit the _Index_ parameter, the new animation behavior is added at the end of the existing animation behaviors.
