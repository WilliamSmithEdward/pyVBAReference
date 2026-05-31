# Interior

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020870-0000-0000-C000-000000000046}  

Represents the interior of an object.

**Example:**

```vba
Worksheets("Sheet1").Range("A1").Interior.ColorIndex = 3
```

## Properties (14)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Color As Variant  (read/write)`  
  Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the RGB function to create a color value. Read/write Variant.
- `ColorIndex As Variant  (read/write)`  
  Returns or sets a Variant value that represents the color of the interior.
- `InvertIfNegative As Variant  (read/write)`  
  True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/write Variant.
- `Pattern As Variant  (read/write)`  
  Returns or sets a Variant value, containing an XlPattern constant, that represents the interior pattern.
- `PatternColor As Variant  (read/write)`  
  Returns or sets the color of the interior pattern as an RGB value. Read/write Variant.
- `PatternColorIndex As Variant  (read/write)`  
  Returns or sets the color of the interior pattern as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Read/write Long.
- `ThemeColor As Variant  (read/write)`  
  Returns or sets a Variant value, containing an XlThemeColor constant, that represents the color. Read/write Variant.
- `TintAndShade As Variant  (read/write)`  
  Returns or sets a Single that lightens or darkens a color.
- `PatternThemeColor As Variant  (read/write)`  
  Returns or sets a theme color pattern for an Interior object. Read/write Variant.
- `PatternTintAndShade As Variant  (read/write)`  
  Returns or sets a tint and shade pattern for an Interior object. Read/write Variant.
- `Gradient As Object  (read-only)`  
  Returns or sets the Gradient property of an Interior object of a selection. Read-only.
