# Interior

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {B184502B-587A-4C6A-8DC4-ECE4354883C6}  

Represents the interior of an object.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 With .Chart.ChartGroups(1)
 .HasUpDownBars = True
 .UpBars.Interior.ColorIndex = 4
 End With
 End If
End With
```

## Properties (9)

- `Color As Variant  (read/write)`  
  Returns or sets the primary color of the object. Read/write Variant.
- `ColorIndex As Variant  (read/write)`  
  Returns or sets the color of the interior. Read/write Variant.
- `InvertIfNegative As Variant  (read/write)`  
  True if Microsoft Word inverts the pattern in the object when it corresponds to a negative number. Read/write Variant.
- `Pattern As Variant  (read/write)`  
  Returns or sets a Variant value, containing an XlPattern constant, that represents the interior pattern.
- `PatternColor As Variant  (read/write)`  
  Returns or sets the color of the interior pattern as an RGB value. Read/write Variant.
- `PatternColorIndex As Variant  (read/write)`  
  Returns or sets the color of the interior pattern as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Read/write Long.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
