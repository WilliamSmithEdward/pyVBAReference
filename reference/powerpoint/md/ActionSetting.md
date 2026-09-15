# ActionSetting

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149348D-5A91-11CF-8700-00AA0060263B}  

Contains information about how the specified shape or text range reacts to mouse actions during a slide show.

**Remarks:** The ActionSetting object is a member of the ActionSettings collection. The ActionSettings collection contains one ActionSetting object that represents how the specified object reacts when the user clicks it during a slide show and one ActionSetting object that represents how the specified object reacts when the user moves the mouse pointer over it during a slide show. If you've set properties of the ActionSetting object that don't seem to be taking effect, make sure that you've set the Action property to the appropriate value.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes(3) _
        .TextFrame.TextRange.ActionSettings(ppMouseClick)
    .Action = ppActionHyperlink
    .Hyperlink.Address = "https://www.microsoft.com"
End With
```

## Properties (10)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Action As PpActionType  (read/write)`  
  Returns or sets the type of action that will occur when the specified shape is clicked or the mouse pointer is positioned over the shape during a slide show. Read/write.
- `ActionVerb As String  (read/write)`  
  Returns or sets a string that contains the OLE verb that will be run when the user clicks the specified shape or passes the mouse pointer over it during a slide show. Read/write.
- `AnimateAction As MsoTriState  (read/write)`  
  Specifies whether the color of the specified shape is momentarily inverted when the specified mouse action occurs. Read/write.
- `Run As String  (read/write)`  
  Returns or sets the name of the presentation or macro to be run when the specified shape is clicked or the mouse pointer passes over the shape during a slide show. Read/write.
- `SlideShowName As String  (read/write)`  
  Returns or sets the name of the custom slide show to run in response to a mouse action on the shape during a slide show. Read/write.
- `Hyperlink As Hyperlink  (read-only)`  
  Returns a Hyperlink object that represents the hyperlink for the specified shape. Read-only.
- `SoundEffect As SoundEffect  (read-only)`  
  Returns a SoundEffect object that represents the sound to be played during the transition to the specified slide. Read-only.
- `ShowAndReturn As MsoTriState  (read/write)`  
  Determines if and under what circumstances Microsoft PowerPoint returns to the initiating slide show. Read/write.
