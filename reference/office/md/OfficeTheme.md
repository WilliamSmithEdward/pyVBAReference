# OfficeTheme

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03A0-0000-0000-C000-000000000046}  

Represents a Microsoft Office theme.

**Example:**

```vba
Dim tTheme As OfficeTheme
Dim tcsThemeColorScheme As ThemeColorScheme
Set tcsThemeColorScheme = tTheme.ThemeColorScheme
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the OfficeTheme object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the OfficeTheme object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the OfficeTheme object. Read-only.
- `ThemeColorScheme As ThemeColorScheme  (read-only)`  
  Gets a ThemeColorScheme object that represents the color scheme of a Microsoft Office theme. Read-only.
- `ThemeFontScheme As ThemeFontScheme  (read-only)`  
  Gets a ThemeFontScheme object that represents the font scheme of a Microsoft Office theme. Read-only.
- `ThemeEffectScheme As ThemeEffectScheme  (read-only)`  
  Gets a ThemeEffectScheme object that represents the effects scheme of a Microsoft Office theme. Read-only.
