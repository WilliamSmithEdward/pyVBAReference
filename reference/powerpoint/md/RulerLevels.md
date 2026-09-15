# RulerLevels

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493491-5A91-11CF-8700-00AA0060263B}  

A collection of all the RulerLevel objects on the specified ruler.

**Remarks:** Each RulerLevel object represents the first-line and left indent for text at a particular outline level. This collection always contains five members - one for each of the available outline levels.

**Example:**

```vba
With ActivePresentation.SlideMaster.TextStyles(ppBodyStyle).Ruler

    .Levels(1).FirstMargin = 0

    .Levels(1).LeftMargin = 40

    .Levels(2).FirstMargin = 60

    .Levels(2).LeftMargin = 100

    .Levels(3).FirstMargin = 120

    .Levels(3).LeftMargin = 160

    .Levels(4).FirstMargin = 180

    .Levels(4).LeftMargin = 220

    .Levels(5).FirstMargin = 240

    .Levels(5).LeftMargin = 280

End With
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (1)

- `Item(Index As Long) As RulerLevel`  
  Returns a single RulerLevel object from the specified RulerLevels collection.
    - `Index As Long` (required): The index number of the single RulerLevel object in the collection to be returned.
