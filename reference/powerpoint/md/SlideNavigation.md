# SlideNavigation

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {F658E3EC-F2D3-4272-AA49-4EC155D5AA76}  

Represents the slide navigation screen in slide show view.

**Example:**

```vba
Sub ShowSlideNavigation()

    ' Start the slide show.
    ActivePresentation.SlideShowSettings.Run

    ' Show the slide navigation screen.
    ActivePresentation.SlideShowWindow.SlideNavigation.Visible = True

End Sub
```

## Properties (3)

- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Visible As Boolean  (read/write)`  
  Specifies whether the slide navigation screen is displayed during a slide show. Read/write.
