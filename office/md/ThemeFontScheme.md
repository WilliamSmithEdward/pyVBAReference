# ThemeFontScheme

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03A5-0000-0000-C000-000000000046}  

Represents the font scheme of a Microsoft Office theme.

**Example:**

```vba
Dim tTheme As OfficeTheme
Dim tfsThemeFontScheme As ThemeFontScheme
Set tfsThemeFontScheme = tTheme.ThemeFontScheme
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ThemeFontScheme object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ThemeFontScheme object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ThemeFontScheme object. Read-only.
- `MinorFont As ThemeFonts  (read-only)`  
  Gets the font settings for the Body of a document. Read-only.
- `MajorFont As ThemeFonts  (read-only)`  
  Gets the font setting for the Headings in a document. Read-only.

## Methods (2)

- `Load(FileName As String)`  
  Loads the font scheme of a Microsoft Office theme from a file.
    - `FileName As String` (required): The name of the font scheme file.
- `Save(FileName As String)`  
  Saves the font scheme of a Microsoft Office theme to a file.
    - `FileName As String` (required): The name of the file.
