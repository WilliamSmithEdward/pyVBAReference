# ThemeFonts

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03A4-0000-0000-C000-000000000046}  

Represents a collection of major and minor fonts in the font scheme of a Microsoft Office theme.

**Example:**

```vba
Dim tTheme As OfficeTheme
Dim tfThemeFonts As ThemeFonts
Set tfThemeFonts = tTheme.ThemeFontScheme.MinorFont
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ThemeFonts object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ThemeFonts object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ThemeFonts object. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the ThemeFonts collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As MsoFontLanguageIndex) As ThemeFont`  
  Gets one of the three language fonts contained in the ThemeFonts collection.
    - `Index As MsoFontLanguageIndex` (required): The index value of the ThemeFont object.
