# ThemeFont

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03A3-0000-0000-C000-000000000046}  

Represents a container for the font schemes of a Microsoft Office theme.

**Example:**

```vba
Dim tTheme As OfficeTheme
Dim tfThemeFontScheme As ThemeFontScheme
Dim tfThemeFont As ThemeFont
Set tfThemeFontScheme = tTheme.ThemeFontScheme
Set tfThemeFont = tfThemeFontScheme.MajorFont(msoThemeLatin)
```

## Properties (4)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ThemeFont object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ThemeFont object was created. Read-only.
- `Name As String  (read/write)`  
  Gets or sets the name of a font in the font scheme of a Microsoft Office theme. Read/write.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ThemeFont object. Read-only.
