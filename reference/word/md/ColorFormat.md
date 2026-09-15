# ColorFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209C6-0000-0000-C000-000000000046}  

Represents the color of a one-color object or the foreground or background color of an object with a gradient or patterned fill. You can set colors to an explicit red-green-blue value by using the RGB property.

**Remarks:** Use one of the properties listed in the following table to return a ColorFormat object. Use the RGB property to set a color to an explicit red-green-blue value. The following example adds a rectangle to the active document and then sets the foreground color, background color, and gradient for the rectangle's fill.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ColorFormat object.
- `RGB As Long  (read/write)`  
  Returns or sets the red-green-blue (RGB) value of the specified color. Read/write Long.
- `Type As MsoColorType  (read-only)`  
  Returns or sets the shape color type. Read-only MsoColorType.
- `TintAndShade As Single  (read/write)`  
  Returns a Single that represents the lightening or darkening of a specified shape's color. Read/write.
- `ObjectThemeColor As WdThemeColorIndex  (read/write)`  
  Returns or sets a WdThemeColorIndex constant that represents the theme color for a color format. Read/write.
- `Brightness As Single  (read/write)`  
  Returns a Single that represents the brightness of a specified shape color. Read/write.
