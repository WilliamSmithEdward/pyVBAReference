# ActionSettings

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149348C-5A91-11CF-8700-00AA0060263B}  

A collection that contains the two ActionSetting objects for a shape or text range. One ActionSetting object represents how the specified object reacts when the user clicks it during a slide show, and the other ActionSetting object represents how the specified object reacts when the user moves the mouse pointer over it during a slide show.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes(3) _
        .ActionSettings(ppMouseOver)
    .Action = ppActionRunMacro
    .Run = "CalculateTotal"
    .AnimateAction = True
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

- `Item(Index As PpMouseActivation) As ActionSetting`  
  Returns a single action setting from the specified ActionSettings collection.
    - `Index As PpMouseActivation` (required): The action setting for a MouseClick or MouseOver event.
