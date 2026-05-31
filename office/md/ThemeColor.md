# ThemeColor

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03A1-0000-0000-C000-000000000046}  

Represents a color in the color scheme of a Microsoft Office theme.

**Example:**

```vba
Dim tcsThemeColorScheme As ThemeColorScheme
Dim tcThemeColor As ThemeColor
Set tcThemeColor = tcsThemeColorScheme.Colors(msoThemeAccent1)
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ThemeColor object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ThemeColor object was created. Read-only.
- `RGB As MsoRGBType  (read/write)`  
  Gets or sets a value of a color in the color scheme of a Microsoft Office theme. Read/write.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ThemeColor object. Read-only.
- `ThemeColorSchemeIndex As MsoThemeColorSchemeIndex  (read-only)`  
  Gets the index value a color scheme of a Microsoft Office theme. Read-only.
