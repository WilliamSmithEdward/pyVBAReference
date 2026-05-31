# ThemeColorScheme

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03A2-0000-0000-C000-000000000046}  

Represents the color scheme of a Microsoft Office theme.

**Example:**

```vba
Dim tTheme As OfficeTheme
Dim tcsThemeColorScheme As ThemeColorScheme
Set tcsThemeColorScheme = tTheme.ThemeColorScheme
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ThemeColorScheme object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ThemeColorScheme object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ThemeColorScheme object. Read-only.
- `Count As Long  (read-only)`  
  Gets an Integer indicating the number of items in the ThemeColorScheme collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (4)

- `Colors(Index As MsoThemeColorSchemeIndex) As ThemeColor`  
  Gets an object that represents a color in the color scheme of a Microsoft Office theme.
    - `Index As MsoThemeColorSchemeIndex` (required): The index value of the ThemeColor object.
- `Load(FileName As String)`  
  Loads the color scheme of a Microsoft Office theme from a file.
    - `FileName As String` (required): The name of the color theme file.
- `Save(FileName As String)`  
  Saves the color scheme of a Microsoft Office theme to a file.
    - `FileName As String` (required): The name of the file.
- `GetCustomColor(Name As String) As MsoRGBType`  
  Gets a value that represents a color in the color scheme of a Microsoft Office theme.
    - `Name As String` (required): The name of the custom color.
