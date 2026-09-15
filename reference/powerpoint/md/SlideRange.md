# SlideRange

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149346B-5A91-11CF-8700-00AA0060263B}  

A collection that represents a notes page or a slide range, which is a set of slides that can contain as little as a single slide or as much as all the slides in a presentation.

**Remarks:** You can include whichever slides you want - chosen from all the slides in the presentation or from all the slides in the selection - to construct a slide range. For example, you could construct a SlideRange collection that contains the first three slides in a presentation, all the selected slides in the presentation, or all the title slides in the presentation. Just as you can work with several slides at the same time in the user interface by selecting them and applying a command, you can work with several slides at the same time programmatically by constructing a SlideRange collection and applying properties or methods to it. And just as some commands in the user interface that work on single slides aren't valid when multiple slides are selected, some properties and methods that work on a Slide object or on a SlideRange collection that contains only one slide will fail if they're applied to a SlideRange collection that contains more than one slide.

**Example:**

```vba
With ActivePresentation.Slides.Range(Array(1, 3))

    .FollowMasterBackground = False
    .Background.Fill.PresetGradient msoGradientHorizontal, _
         1, msoGradientLateSunset

End With
```

## Properties (29)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Shapes As Shapes  (read-only)`  
  Returns a Shapes collection that represents all the elements that have been placed or inserted on the specified slide, slide master, or range of slides. Read-only.
- `HeadersFooters As HeadersFooters  (read-only)`  
  Returns a HeadersFooters collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides. Read-only.
- `SlideShowTransition As SlideShowTransition  (read-only)`  
  Returns a SlideShowTransition object that represents the special effects for the specified slide transition. Read-only.
- `ColorScheme As ColorScheme  (read/write)`  
  Returns or sets the ColorScheme object that represents the scheme colors for the specified slide, slide range, or slide master. Read/write.
- `Background As ShapeRange  (read-only)`  
  Returns a ShapeRange object that represents the slide background.
- `Name As String  (read/write)`  
  When a slide is inserted into a presentation, Microsoft PowerPoint automatically assigns it a name in the form Slide _n_, where _n_ is an integer that represents the order in which the slide was created in the presentation. For example, the first slide inserted into a presentation is automatically named Slide1. If you copy a slide from one presentation to another, the slide loses the name it had in the first presentation and is automatically assigned a new name in the second presentation. A slide range must contain exactly one slide. Read/write.
- `SlideID As Long  (read-only)`  
  Returns a unique ID number for the specified slide. Read-only.
- `PrintSteps As Long  (read-only)`  
  Returns the number of slides you'd need to print to simulate the builds on the specified slide, slide master, or range of slides. Read-only.
- `Layout As PpSlideLayout  (read/write)`  
  Returns or sets a PpSlideLayout constant that represents the slide layout. Read/write.
- `Tags As Tags  (read-only)`  
  Returns a Tags object that represents the tags for the specified object. Read-only.
- `SlideIndex As Long  (read-only)`  
  Returns the index number of the specified slide within the Slides collection. Read-only.
- `SlideNumber As Long  (read-only)`  
  Returns the slide number. Read-only.
- `DisplayMasterShapes As MsoTriState  (read/write)`  
  Determines whether the specified range of slides displays the background objects on the slide master. Read/write.
- `FollowMasterBackground As MsoTriState  (read/write)`  
  Determines whether the range of slides follows the slide master background. Read/write.
- `NotesPage As SlideRange  (read-only)`  
  Returns a SlideRange object that represents the notes pages for the specified slide or range of slides. Read-only.
- `Master As _Master  (read-only)`  
  Returns a Master object that represents the slide master. Read-only.
- `Hyperlinks As Hyperlinks  (read-only)`  
  Returns a Hyperlinks collection that represents all the hyperlinks on the specified slide. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Comments As Comments  (read-only)`  
  Returns a Comments object that represents a collection of comments. Read-only.
- `Design As Design  (read/write)`  
  Returns a Design object representing a design.
- `TimeLine As TimeLine  (read-only)`  
  Returns a TimeLine object that represents the animation timeline for the slide. Read-only.
- `CustomLayout As CustomLayout  (read/write)`  
  Returns a CustomLayout object that represents the custom layout associated with the specified range of slides. Read-only.
- `ThemeColorScheme As ThemeColorScheme  (read-only)`  
  Returns a ThemeColorScheme object that represents the color scheme associated with the specified range of slides. Read-only.
- `BackgroundStyle As MsoBackgroundStyleIndex  (read/write)`  
  Sets or returns the background style of the specified object. Read/write.
- `CustomerData As CustomerData  (read-only)`  
  Returns a CustomerData object. Read-only.
- `sectionIndex As Long  (read-only)`  
  Returns the index of the selected section in the SlideRange. Read-only.
- `HasNotesPage As MsoTriState  (read-only)`  
  Indicates whether the selected SlideRange has media that resides on a notes page. Read-only.

## Methods (13)

- `Select()`  
  Selects the specified object.
- `Cut()`  
  Deletes the specified object and places it on the Clipboard.
- `Copy()`  
  Copies the specified object to the Clipboard.
- `Duplicate() As SlideRange`  
  Creates a duplicate of the specified SlideRange object, adds the new range of slides to the Slides collection immediately after the slide range specified originally, and then returns a SlideRange object that represents the duplicate slides.
- `Delete()`  
  Deletes the specified SlideRange object.
- `Export(FileName As String, FilterName As String, [ScaleWidth As Long], [ScaleHeight As Long])`  
  Exports a range of slides, using the specified graphics filter, and saves the exported file under the specified file name.
    - `FileName As String` (required): The name of the file to be exported and saved to disk. You can include a full path; if you don't, Microsoft PowerPoint creates a file in the current folder.
    - `FilterName As String` (required): The graphics format in which you want to export slides. The specified graphics format must have an export filter registered in the Windows registry. You can specify either the registered extension or the registered filter name. Microsoft PowerPoint will first search for a matching extension in the registry. If no extension that matches the specified string is found, PowerPoint will look for a filter name that matches.
    - `ScaleWidth As Long` (optional): The width in pixels of an exported slide.
    - `ScaleHeight As Long` (optional): The height in pixels of an exported slide.
- `Item(Index As Variant) As Slide`  
  Returns a single Slide object from the specified SlideRange collection.
    - `Index As Variant` (required): The name or index number of the single Slide object in the collection to be returned.
- `MoveTo(toPos As Long)`  
  Moves the specified object to a specific location within the same collection, renumbering all other items in the collection appropriately.
    - `toPos As Long` (required): The index position to which to move the animation effect.
- `ApplyTemplate(FileName As String)`  
  Applies a design template to the specified slide range.
    - `FileName As String` (required): Specifies the name of the design template.
- `ApplyTheme(themeName As String)`  
  Applies a theme or design template to the specified range of slides.
    - `themeName As String` (required): The path and name of the theme file (.thmx) or design template file (.pot) to apply to the SlideRange object.
- `ApplyThemeColorScheme(themeColorSchemeName As String)`  
  Applies a color scheme to the specified range of slides.
    - `themeColorSchemeName As String` (required): The path and name of the color scheme file to apply to the range of slides.
- `MoveToSectionStart(toSection As Long)`  
  Moves the current position to the start of the specified section in the SlideRange object.
    - `toSection As Long` (required): The section to move to.
- `ApplyTemplate2(FileName As String, VariantGUID As String)`  
  Applies a design template and theme variant to the slide range.
