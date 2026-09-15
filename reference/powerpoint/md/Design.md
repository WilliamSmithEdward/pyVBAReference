# Design

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934D7-5A91-11CF-8700-00AA0060263B}  

Represents an individual slide design template. The Design object is a member of the Designs and SlideRange collections and the Master and Slide objects.

**Remarks:** Use the Design property of the Master, Slide, or SlideRange objects to access a Design object, for example: - ActivePresentation.SlideMaster.Design - ActivePresentation.Slides(1).Design - ActivePresentation.Slides.Range.Design Use the Add, Item, Clone, or Loadmethods of the Designs collection to add, refer to, clone, or load a Design object, respectively. For example, to add a design template, use ActivePresentation.Designs.Add designName:="MyDesign"

**Example:**

```vba
Sub AddQueryTitleMaster(dsn As Design)

    dsn.AddTitleMaster

    MsgBox dsn.HasTitleMaster

End Sub
```

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `SlideMaster As _Master  (read-only)`  
  Returns a Master object that represents the slide master. Read-only.
- `Index As Long  (read-only)`  
  Returns a Long that represents the index number for an animation effect or design. Read-only.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write.
- `Preserved As MsoTriState  (read/write)`  
  Represents whether a design master is preserved from changes. Read/write.

## Methods (2)

- `MoveTo(toPos As Long)`  
  Moves the specified object to a specific location within the same collection, renumbering all other items in the collection appropriately.
    - `toPos As Long` (required): The index position to which to move the animation effect.
- `Delete()`  
  Deletes the specified Design object.
