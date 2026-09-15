# RGBColor

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493470-5A91-11CF-8700-00AA0060263B}  

Represents a single color in a color scheme.

**Example:**

```vba
With ActivePresentation.ColorSchemes

    .Item(1).Colors(ppBackground).RGB = RGB(255, 0, 0)

    .Item(1).Colors(ppTitle) = .Item(2).Colors(ppTitle)

End With
```

## Properties (3)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `RGB As MsoRGBType  (read/write)`  
  Returns or sets the red-green-blue (RGB) value of a specified color-scheme color or extra color when used with a PpColorSchemeIndex constant. Read/write.
