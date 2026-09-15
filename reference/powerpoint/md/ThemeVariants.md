# ThemeVariants

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9E116A3C-2C6D-4D07-93AF-8675D452FCA2}  

A collection of ThemeVariant objects that represent variations in the theme.

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

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.

## Methods (1)

- `Item(Index As Long) As ThemeVariant`  
  Returns a single ThemeVariant object from the collection.
