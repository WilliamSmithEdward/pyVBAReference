# Sequences

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934DD-5A91-11CF-8700-00AA0060263B}  

Represents a collection of Sequence objects. Use a Sequence object to add, find, modify, and clone animation effects.

**Example:**

```vba
Sub AddNewSequence()

    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim interEffect As Effect

    Set shp1 = ActivePresentation.Slides(1).Shapes.AddShape _
        (Type:=msoShape32pointStar, Left:=100, _
        Top:=100, Width:=200, Height:=200)

    Set shp2 = ActivePresentation.Slides(1).Shapes.AddShape _
        (Type:=msoShapeBevel, Left:=400, _
        Top:=200, Width:=150, Height:=100)

    With ActivePresentation.Slides(1).TimeLine.InteractiveSequences.Add(1)
        Set interEffect = .AddEffect(shp2, msoAnimEffectBlinds, _
            trigger:=msoAnimTriggerOnShapeClick)
        interEffect.Shape = shp1
    End With

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

- `Item(Index As Long) As Sequence`  
  Returns a single Sequence object from the specified Sequences collection.
    - `Index As Long` (required): The index number of the single Sequence object in the collection to be returned.
- `Add([Index As Long]) As Sequence`  
  Returns a Sequence object that represents a new sequence.
    - `Index As Long` (optional): The position of the sequence in relation to other sequences. The default value is -1, which means that if you omit the Index parameter, the new sequence is added to the end of the existing sequences.
