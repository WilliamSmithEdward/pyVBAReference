# ChartFont

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {CDB0FF41-E862-47BB-AE77-3FA7B1AE3189}  

Contains the font attributes (font name, font size, color, and so on) for an object chart.

**Remarks:** If you don't want to format all the text in an AxisTitle, ChartTitle, DataLabel, or DisplayUnitLabel object the same way, use the Characters property of that object to first return a subset of the text as a ChartCharacters object. Then use the Font property of the ChartCharacters object to return a ChartFont object you can use to format the subset of text, as needed.

**Example:**

```vba
With ActiveDocument.InlineShapes(1).Chart
 .AxisTitle.Font.Bold = True
End With
```

## Properties (15)

- `Background As Variant  (read/write)`  
  Returns or sets the type of background for text used in charts. Read/write Variant that is set to one of the constants of XlBackground.
- `Bold As Variant  (read/write)`  
  True if the font is bold. Read/write Variant.
- `Color As Variant  (read/write)`  
  Returns or sets the primary color of the object. Read/write Variant.
- `ColorIndex As Variant  (read/write)`  
  Returns or sets the color of the font. Read/write Variant.
- `FontStyle As Variant  (read/write)`  
  Returns or sets the font style. Read/write String.
- `Italic As Variant  (read/write)`  
  True if the font style is italic. Read/write Boolean.
- `Name As Variant  (read/write)`  
  Returns or sets the name of the object. Read/write String.
- `Size As Variant  (read/write)`  
  Returns or sets the size of the font. Read/write Variant.
- `StrikeThrough As Variant  (read/write)`  
  True if the font is struck through with a horizontal line. Read/write Boolean.
- `Subscript As Variant  (read/write)`  
  True if the font is formatted as subscript. The default is False. Read/write Variant.
- `Superscript As Variant  (read/write)`  
  True if the font is formatted as superscript. The default is False. Read/write Variant.
- `Underline As Variant  (read/write)`  
  Returns or sets the type of underline applied to the font. Can be one of the XlUnderlineStyle constants. Read/write Variant.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
