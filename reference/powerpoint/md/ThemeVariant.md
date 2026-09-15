# ThemeVariant

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {C9195677-B4F9-4228-BFD0-40C1F77D2F6A}  

Represents a variation (set of colors and fonts) in a theme.

**Example:**

```vba
Sub IterateThemeVariants()

    Dim pptTheme As Theme
    Dim pptThemeVariants As ThemeVariants
    Dim pptThemeVariant As ThemeVariant
    Dim path As String

    ' Get a reference to the currently active theme.
    path = "C:\Program Files (x86)\Microsoft Office\Document Themes 15\" & _
        ActivePresentation.TemplateName & ".thmx"
    Set pptTheme = Application.OpenThemeFile(path)

    ' Get a reference to all of the variations in the theme.
    Set pptThemeVariants = pptTheme.ThemeVariants

    ' Iterate over each variation of the theme and print
    ' its name and ID.
    For Each pptThemeVariant In pptThemeVariants

        Debug.Print "Variation " & pptThemeVariant.name & " id: " & pptThemeVariant.Id

    Next pptThemeVariant

End Sub
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Name As String  (read-only)`  
  Returns the name of the theme variation. Read-only.
- `Width As Long  (read-only)`  
  Returns the height of the theme variation. Read-only.
- `Height As Long  (read-only)`  
  Returns the height of the theme variation. Read-only.
- `Id As String  (read-only)`  
  Returns a string that represents the ID of the theme variation. Read-only.
