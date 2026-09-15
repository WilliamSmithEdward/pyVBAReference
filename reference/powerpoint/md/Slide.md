# Slide

**Type:** Class  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493445-5A91-11CF-8700-00AA0060263B}  

Represents a slide. The Slides collection contains all the Slide objects in a presentation.

**Remarks:** The following examples describe how to: - Return a slide that you specify by name, index number, or slide ID number - Return a slide in the selection - Return the slide that's currently displayed in any document window or slide show window you specify - Create a new slide

**Example:**

```vba
ActivePresentation.Slides(1).Layout = ppLayoutTitle
```

## Properties (28)

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
  When a slide is inserted into a presentation, Microsoft PowerPoint automatically assigns it a name in the form Slide _n_, where _n_ is an integer that represents the order in which the slide was created in the presentation.
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
  Determines whether the specified slide displays the background objects on the slide master. Read/write.
- `FollowMasterBackground As MsoTriState  (read/write)`  
  Determines whether the slide follows the slide master background. Read/write.
- `NotesPage As SlideRange  (read-only)`  
  Returns a SlideRange object that represents the notes pages for the specified slide or range of slides. Read-only.
- `Master As _Master  (read-only)`  
  Returns a Master object that represents the slide master. Read-only.
- `Hyperlinks As Hyperlinks  (read-only)`  
  Returns a Hyperlinks collection that represents all the hyperlinks on the specified slide. Read-only.
- `Comments As Comments  (read-only)`  
  Returns a Comments object that represents a collection of comments. Read-only.
- `Design As Design  (read/write)`  
  Returns a Design object representing a design.
- `TimeLine As TimeLine  (read-only)`  
  Returns a TimeLine object that represents the animation timeline for the slide. Read-only.
- `CustomLayout As CustomLayout  (read/write)`  
  Returns a CustomLayout object that represents the custom layout associated with the specified slide. Read-only.
- `ThemeColorScheme As ThemeColorScheme  (read-only)`  
  Returns a ThemeColorScheme object that represents the color scheme associated with the specified slide. Read-only.
- `BackgroundStyle As MsoBackgroundStyleIndex  (read/write)`  
  Sets or returns the background style of the specified object. Read/write.
- `CustomerData As CustomerData  (read-only)`  
  Returns a CustomerData object. Read-only.
- `sectionIndex As Long  (read-only)`  
  Returns the index of the selected section in the Slide range. Read-only.
- `HasNotesPage As MsoTriState  (read-only)`  
  Indicates whether the selected Slide has media that resides on a notes page. Read-only.

## Methods (12)

- `Select()`  
  Selects the specified object.
- `Cut()`  
  Deletes the specified object and places it on the Clipboard.
- `Copy()`  
  Copies the specified object to the Clipboard.
- `Duplicate() As SlideRange`  
  Creates a duplicate of the specified Slide object, adds the new slide to the Slides collection immediately after the slide specified originally, and then returns a Slide object that represents the duplicate slide.
- `Delete()`  
  Deletes the specified Slide object.
- `Export(FileName As String, FilterName As String, [ScaleWidth As Long], [ScaleHeight As Long])`  
  Exports a slide, using the specified graphics filter, and saves the exported file under the specified file name.
    - `FileName As String` (required): The name of the file to be exported and saved to disk. You can include a full path; if you don't, Microsoft PowerPoint creates a file in the current folder.
    - `FilterName As String` (required): The graphics format in which you want to export slides. The specified graphics format must have an export filter registered in the Windows registry. You can specify either the registered extension or the registered filter name. Microsoft PowerPoint will first search for a matching extension in the registry. If no extension that matches the specified string is found, PowerPoint will look for a filter name that matches.
    - `ScaleWidth As Long` (optional): The width in pixels of an exported slide.
    - `ScaleHeight As Long` (optional): The height in pixels of an exported slide.
- `MoveTo(toPos As Long)`  
  Moves the specified object to a specific location within the same collection, renumbering all other items in the collection appropriately.
    - `toPos As Long` (required): The index position to which to move the animation effect.
- `ApplyTemplate(FileName As String)`  
  Applies a design template to the specified slide.
    - `FileName As String` (required): Specifies the name of the design template.
- `ApplyTheme(themeName As String)`  
  Applies a theme or design template to the specified slide.
    - `themeName As String` (required): The path and name of the theme file (.thmx) or design template file (.pot) to apply to the Slide object.
- `ApplyThemeColorScheme(themeColorSchemeName As String)`  
  Applies a color scheme to the specified slide.
    - `themeColorSchemeName As String` (required): The path and name of the color scheme file to apply to the slide.
- `MoveToSectionStart(toSection As Long)`  
  Moves the current slide to the start of the specified section.
    - `toSection As Long` (required): The section to move to.
- `ApplyTemplate2(FileName As String, VariantGUID As String)`  
  Applies a design template and theme variant to the slide.
