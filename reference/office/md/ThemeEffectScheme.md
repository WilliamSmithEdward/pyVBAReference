# ThemeEffectScheme

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03A6-0000-0000-C000-000000000046}  

Represents the effects scheme of a Microsoft Office theme.

**Example:**

```vba
Dim tesEffectScheme As ThemeEffectScheme
tesEffectScheme.Load("C:\myThemeEffectScheme.eftx")
```

## Properties (3)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ThemeEffectScheme object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ThemeEffectScheme object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ThemeEffectScheme object. Read-only.

## Methods (1)

- `Load(FileName As String)`  
  Loads the effects scheme of a Microsoft Office theme from a file.
    - `FileName As String` (required): The name of the effect scheme file.
