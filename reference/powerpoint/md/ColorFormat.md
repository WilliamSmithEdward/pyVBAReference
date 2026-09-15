# ColorFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493452-5A91-11CF-8700-00AA0060263B}  

Represents the color of a one-color object, the foreground or background color of an object with a gradient or patterned fill, or the pointer color. You can set colors to an explicit red-green-blue value (by using the RGB property) or to a color in the theme (by using the [ObjectThemeColor]) or the legacy pre-Office 2007 color scheme (by using the SchemeColor property).

**Remarks:** Use one of the properties listed in the following table to return a ColorFormat object.

**Example:**

```vba
ActivePresentation.Slides(2).Shapes(1).TextFrame.TextRange.Font.Color.SchemeColor = ppTitle
```

## Properties (9)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `RGB As MsoRGBType  (read/write)`  
  Returns or sets the red-green-blue (RGB) value of the specified color. Read/write.
- `Type As MsoColorType  (read-only)`  
  Represents the type of color. Read-only.
- `SchemeColor As PpColorSchemeIndex  (read/write)`  
  Returns or sets the color in the applied legacy color scheme that's associated with the specified object. Provided for backwards compatibility. Use the [ObjectThemeColor] instead for Office 2007 and later. Read/write.
- `TintAndShade As Single  (read/write)`  
  Sets or returns the lightening or darkening of the color of a specified shape. Read/write.
- `ObjectThemeColor As MsoThemeColorIndex  (read/write)`  
  Returns or sets the theme color of the specified ColorFormat object. Read/write.
- `Brightness As Single  (read/write)`  
  Returns or sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write Single.
