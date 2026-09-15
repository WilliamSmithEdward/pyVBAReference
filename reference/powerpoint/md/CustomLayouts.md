# CustomLayouts

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934F2-5A91-11CF-8700-00AA0060263B}  

Represents a set of custom layouts associated with a presentation design.

**Remarks:** Use the CustomLayouts property of the slide Master object to return a CustomLayouts collection. Use CustomLayouts (index), where index is the custom layout index number, to return a single CustomLayout object. Use the Add method to create a new custom layout and add it to the CustomLayouts collection. Use the Paste method to past slides from the Clipboard as a CustomLayout object into the CustomLayouts collection. Use the CustomLayout property of a Slide or SlideRange object to return a custom layout for a slide or set of slides.

**Example:**

```vba
Sub AddCustomLayout()

    With ActivePresentation.SlideMaster

        .CustomLayouts.Add (1)

        .CustomLayouts(1).Name = "MyLayout"

    End With

End Sub
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.

## Methods (3)

- `Item(Index As Variant) As CustomLayout`  
  Returns a single object from the specified collection.
    - `Index As Variant` (required): The name or index number of the single object in the collection to be returned.
- `Add(Index As Long) As CustomLayout`  
  Returns a CustomLayout object that represents a new custom layout.
    - `Index As Long` (required): The index number of the custom layout. The default value is -1, which means that if the Index argument is omitted, the new custom layout is added at the end of existing custom layouts.
- `Paste([Index As Long]) As CustomLayout`  
  Pastes the slides on the Clipboard into a custom layout and adds the custom layout to the CustomLayouts collection.
    - `Index As Long` (optional): The index number of the custom layout before which the new custom layout is pasted. If this argument is omitted, the new custom layout is pasted at the end of the CustomLayouts collection.
