# Master

**Type:** Class  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493447-5A91-11CF-8700-00AA0060263B}  

Represents a slide master, title master, handout master, notes master, or design master.

**Example:**

```vba
ActivePresentation.SlideMaster.Background.Fill _

    .PresetGradient msoGradientHorizontal, 1, msoGradientBrass
```

## Properties (19)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Shapes As Shapes  (read-only)`  
  Returns a Shapes collection that represents all the elements that have been placed or inserted on the specified slide, slide master, or range of slides. Read-only.
- `HeadersFooters As HeadersFooters  (read-only)`  
  Returns a HeadersFooters collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides. Read-only.
- `ColorScheme As ColorScheme  (read/write)`  
  Returns or sets the ColorScheme object that represents the scheme colors for the specified slide, slide range, or slide master. Read/write.
- `Background As ShapeRange  (read-only)`  
  Returns a ShapeRange object that represents the slide background.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write.
- `Height As Single  (read-only)`  
  Returns or sets the height of the specified object, in points. Read-only.
- `Width As Single  (read-only)`  
  Returns the width of the specified object, in points. Read-only.
- `TextStyles As TextStyles  (read-only)`  
  Returns a TextStyles collection that represents three text styles - title text, body text, and default text - for the specified slide master. Read-only.
- `Hyperlinks As Hyperlinks  (read-only)`  
  Returns a Hyperlinks collection that represents all the hyperlinks on the specified slide. Read-only.
- `Design As Design  (read-only)`  
  Returns a Design object representing a design.
- `TimeLine As TimeLine  (read-only)`  
  Returns a TimeLine object that represents the animation timeline for the slide. Read-only.
- `SlideShowTransition As SlideShowTransition  (read-only)`  
  Returns a SlideShowTransition object that represents the special effects for the specified slide transition. Read-only.
- `CustomLayouts As CustomLayouts  (read-only)`  
  Returns a CustomLayouts object that represents the custom layouts associated with the presentation design of the specified Master object. Read-only.
- `Theme As OfficeTheme  (read-only)`  
  Returns a Theme object that represents the theme used by the specified slide master, title master, handout master, notes master, or design master. Read-only.
- `BackgroundStyle As MsoBackgroundStyleIndex  (read/write)`  
  Sets or returns the background style of the specified object. Read/write.
- `CustomerData As CustomerData  (read-only)`  
  Returns a CustomerData object. Read-only.
- `Guides As Guides  (read-only)`  
  Returns a Guides collection that represents all of the drawing guides associated with the slide master, title master, handout master, notes master, or design master.

## Methods (2)

- `Delete()`  
  Deletes the specified Master object.
- `ApplyTheme(themeName As String)`  
  Applies a theme or design template to the specified slide master, title master, handout master, notes master, or design master.
    - `themeName As String` (required): The path and name of the theme file (.thmx) or design template file (.pot) to apply to the Master object.
