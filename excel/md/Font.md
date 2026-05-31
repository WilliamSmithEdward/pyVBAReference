# Font

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002084D-0000-0000-C000-000000000046}  

Contains the font attributes (font name, font size, color, and so on) for an object.

**Remarks:** If you don't want to format all the text in a cell or graphic the same way, use the Characters property of the Range object to return a subset of the text.

**Example:**

```vba
Worksheets("Sheet1").Range("A1:C5").Font.Bold = True
```

## Properties (18)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Background As Variant  (read/write)`  
  Returns or sets the type of background for text used in charts. Read/write Variant, which is set to one of the constants of XlBackground.
- `Bold As Variant  (read/write)`  
  True if the font is bold. Read/write Variant.
- `Color As Variant  (read/write)`  
  Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the RGB function to create a color value. Read/write Variant.
- `ColorIndex As Variant  (read/write)`  
  Returns or sets a Variant value that represents the color of the font.
- `FontStyle As Variant  (read/write)`  
  Returns or sets the font style. Read/write String.
- `Italic As Variant  (read/write)`  
  True if the font style is italic. Read/write Boolean.
- `Name As Variant  (read/write)`  
  Returns or sets a Variant value that represents the name of the object.
- `Size As Variant  (read/write)`  
  Returns or sets the size of the font. Read/write Variant.
- `Strikethrough As Variant  (read/write)`  
  True if the font is struck through with a horizontal line. Read/write Boolean.
- `Subscript As Variant  (read/write)`  
  True if the font is formatted as subscript. False by default. Read/write Variant.
- `Superscript As Variant  (read/write)`  
  True if the font is formatted as superscript; False by default. Read/write Variant.
- `Underline As Variant  (read/write)`  
  Returns or sets the type of underline applied to the font. Read/write Variant.
- `ThemeColor As Variant  (read/write)`  
  Returns or sets the theme color in the applied color scheme that is associated with the specified object. Read/write Variant.
- `TintAndShade As Variant  (read/write)`  
  Returns or sets a Single that lightens or darkens a color.
- `ThemeFont As XlThemeFont  (read/write)`  
  Returns or sets the theme font in the applied font scheme that is associated with the specified object. Read/write XlThemeFont.
