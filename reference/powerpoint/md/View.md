# View

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493458-5A91-11CF-8700-00AA0060263B}  

Represents the current editing view in the specified document window.

**Example:**

```vba
With Windows(1)

    .Height = 200

    .Width = 250

    .View.ZoomToFit = True

End With
```

## Properties (13)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Type As PpViewType  (read-only)`  
  Represents the type of view. Read-only.
- `Zoom As Long  (read/write)`  
  Returns or sets the zoom setting of the specified view as a percentage of normal size. Read/write.
- `Slide As Object  (read/write)`  
  Returns or sets a Slide object that represents the slide that's currently displayed in the specified document window view. Read/write.
- `DisplaySlideMiniature As MsoTriState  (read/write)`  
  Determines if and when the slide miniature window is displayed automatically. Read/write.
- `ZoomToFit As MsoTriState  (read/write)`  
  Determines whether the view is zoomed to fit the dimensions of the document window every time the document window is resized. Read/write.
- `PrintOptions As PrintOptions  (read-only)`  
  Returns a PrintOptions object that represents print options that are saved with the specified presentation. Read-only.
- `MediaControlsVisible As MsoTriState  (read-only)`  
  Indicates whether the media controls are currently visible. Read-only.
- `MediaControlsLeft As Single  (read-only)`  
  Returns the distance, in points, from the left edge of the media control bounding box to the left edge of the slide. Read-only.
- `MediaControlsTop As Single  (read-only)`  
  Returns the distance, in points, from the top edge of the media control bounding box to the top edge of the slide. Read-only.
- `MediaControlsWidth As Single  (read-only)`  
  Returns the width, in points, of the media control bounding box. Read-only.
- `MediaControlsHeight As Single  (read-only)`  
  Returns the height, in points, of the media controls. Read-only.

## Methods (5)

- `Paste()`  
  Pastes the contents of the Clipboard into the active view. Attempting to paste an object into a view that won't accept it causes an error.
- `GotoSlide(Index As Long)`  
  Switches to the specified slide.
    - `Index As Long` (required): The number of the slide to switch to.
- `PasteSpecial([DataType As PpPasteDataType], [DisplayAsIcon As MsoTriState], [IconFileName As String], [IconIndex As Long], [IconLabel As String], [Link As MsoTriState])`  
  Pastes the current contents of the Clipboard into the view represented by the View object.
    - `DataType As PpPasteDataType` (optional): A format for the Clipboard contents when they're inserted into the document. The default value varies, depending on the contents in the Clipboard. An error occurs if the specified data type in the DataType argument is not supported by the clipboard contents.
    - `DisplayAsIcon As MsoTriState` (optional): msoTrue to display the embedded object (or link) as an icon.
    - `IconFileName As String` (optional): If DisplayAsIcon is set to msoTrue, this argument is the path and file name for the file in which the icon to be displayed is stored. If DisplayAsIcon is set to msoFalse, this argument is ignored.
    - `IconIndex As Long` (optional): If DisplayAsIcon is set to msoTrue, this argument is a number that corresponds to the icon you want to use in the program file specified by IconFilename. Icons appear in the Change Icon dialog box, accessed from the Insert tab (click Object, select Display as icon, click Change Icon): 0 (zero) corresponds to the first icon, 1 corresponds to the second icon. If this argument is omitted, the first (default) icon is used. If DisplayAsIcon is set to msoFalse, this argument is ignored. If IconIndex is outside the valid range, the default icon (index 0) is used.
    - `IconLabel As String` (optional): If DisplayAsIcon is set to msoTrue, this argument is the text that appears below the icon. If this label is missing, Microsoft PowerPoint generates an icon label based on the Clipboard contents. If DisplayAsIcon is set to msoFalse, this argument is ignored.
    - `Link As MsoTriState` (optional): Determines whether to create a link to the source file of the Clipboard contents. An error occurs if the Clipboard contents don't support a link.
- `PrintOut([From As Long], [To As Long], [PrintToFile As String], [Copies As Long], [Collate As MsoTriState])`  
  Prints the specified presentation.
    - `From As Long` (optional): The number of the first page to be printed. If this argument is omitted, printing starts at the beginning of the presentation. Specifying the To and From arguments sets the contents of the PrintRanges object and sets the value of the RangeType property for the presentation.
    - `To As Long` (optional): The number of the last page to be printed. If this argument is omitted, printing continues to the end of the presentation. Specifying the To and From arguments sets the contents of the PrintRanges object and sets the value of the RangeType property for the presentation.
    - `PrintToFile As String` (optional): The name of the file to print to. If you specify this argument, the file is printed to a file rather than sent to a printer. If this argument is omitted, the file is sent to a printer.
    - `Copies As Long` (optional): The number of copies to be printed. If this argument is omitted, only one copy is printed. Specifying this argument sets the value of the NumberOfCopies property.
    - `Collate As MsoTriState` (optional): If this argument is omitted, multiple copies are collated. Specifying this argument sets the value of the Collate property.
- `Player(ShapeId As Variant) As Player`  
  Allows access to playback controls for the associated view in the current window.
    - `ShapeId As Variant` (required): The playback control.
