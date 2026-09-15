# CustomLayout

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934F3-5A91-11CF-8700-00AA0060263B}  

Represents a custom layout associated with a presentation design. The CustomLayout object is a member of the CustomLayouts collection.

**Remarks:** Use the CustomLayout property of the Slide or SlideRange objects to access a CustomLayout object, for example: Use the Add method of the CustomLayouts collection to add a new custom layout to the presentation design's custom layouts. Use the Item method to refer to a custom layout. Use the Paste method to paste the slides on the Clipboard into a custom layout and add the custom layout to the CustomLayouts collection.

## Properties (20)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object. Read-only.
- `Parent As Object  (read-only)`  
  Returns a reference to the slide master object that is the parent of the specified CustomLayout object. Read-only.
- `Shapes As Shapes  (read-only)`  
  Returns a Shapes collection representing all the layout elements included in the specified custom layout. Read-only.
- `HeadersFooters As HeadersFooters  (read-only)`  
  Returns a HeadersFooters collection that represents the header, footer, date and time, and slide number associated with the specified custom layout.
- `Background As ShapeRange  (read-only)`  
  Returns a ShapeRange object that represents the slide background. Read-only.
- `Name As String  (read/write)`  
  Represents the display name of the specified custom layout. Read/write.
- `Height As Single  (read-only)`  
  Returns the height of the specified object, in points. Read-only.
- `Width As Single  (read-only)`  
  Returns the width of the specified custom layout, in points. Read-only.
- `Hyperlinks As Hyperlinks  (read-only)`  
  Returns a Hyperlinks collection that represents all hyperlinks on the slide associated with the specified custom layout. Read-only.
- `Design As Design  (read-only)`  
  Returns a Design object that represents the design template associated with the specified custom layout.
- `TimeLine As TimeLine  (read-only)`  
  Returns a TimeLine object that represents the animation timeline for the slide associated with the specified custom layout. Read-only.
- `SlideShowTransition As SlideShowTransition  (read-only)`  
  Returns a SlideShowTransition object that represents the special effects for the specified custom layout. Read-only.
- `MatchingName As String  (read/write)`  
  Represents the internal name of the specified custom layout. Read/write.
- `Preserved As MsoTriState  (read/write)`  
  Determines whether the specified custom layout is preserved from changes. Read/write.
- `Index As Long  (read-only)`  
  Returns the index position of the specified custom layout in the CustomLayouts collection. Read-only.
- `DisplayMasterShapes As MsoTriState  (read/write)`  
  Determines whether the specified custom layout displays background objects on the slide master. Read/write.
- `FollowMasterBackground As MsoTriState  (read/write)`  
  Determines whether the specified custom layout follows the slide master background. Read/write.
- `ThemeColorScheme As ThemeColorScheme  (read-only)`  
  Returns a ThemeColorScheme object that represents the color scheme of the theme associated with the specified custom layout. Read-only.
- `CustomerData As CustomerData  (read-only)`  
  Returns a CustomerData object.
- `Guides As Guides  (read-only)`  
  Returns the Guides collection associated with a custom layout. Read-only.

## Methods (6)

- `Delete()`  
  Deletes the specified object.
- `Select()`  
  Selects the specified CustomLayout object.
- `Cut()`  
  Deletes the specified object and places it on the Clipboard.
- `Copy()`  
  Copies the specified object to the Clipboard.
- `Duplicate() As CustomLayout`  
  Creates a duplicate of the specified custom layout, adds the new custom layout to the CustomLayouts collection immediately after the original custom layout, and returns a CustomLayout object that represents the duplicate layout.
- `MoveTo(toPos As Long)`  
  Moves the specified CustomLayout object to a different position in the CustomLayouts collection.
    - `toPos As Long` (required): The index position in the CustomLayouts collection to which the CustomLayout object will be moved.
