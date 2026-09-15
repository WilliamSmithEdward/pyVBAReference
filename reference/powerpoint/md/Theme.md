# Theme

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {D9D60EB3-D4B4-4991-9C16-75585B3346BB}  

Represents a theme (a collection of colors, fonts, and effects).

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
    ' its ID.
    For Each pptThemeVariant In pptThemeVariants

        Debug.Print "Variation id: " & pptThemeVariant.Id

    Next pptThemeVariant

End Sub
```

## Properties (3)

- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `ThemeVariants As ThemeVariants  (read-only)`  
  Returns a ThemeVariants collection that represents the variations in the theme.
