# TimeLine

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934DC-5A91-11CF-8700-00AA0060263B}  

Stores animation information for a Master, Slide, or SlideRange object.

**Example:**

```vba
ActivePresentation.Slides(1).TimeLine.MainSequence

ActivePresentation.SlideMaster.TimeLine.InteractiveSequences

ActiveWindow.Selection.SlideRange.TimeLine.InteractiveSequences
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `MainSequence As Sequence  (read-only)`  
  Returns a Sequence object that represents the collection of Effect objects in the main animation sequence of a slide.
- `InteractiveSequences As Sequences  (read-only)`  
  Returns a Sequences object that represents animations that are triggered by click a specified shape.
