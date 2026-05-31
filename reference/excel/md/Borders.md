# Borders

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020855-0000-0000-C000-000000000046}  

A collection of four Border objects that represent the four borders of a Range object or Style object.

**Remarks:** Use the Borders property to return the Borders collection, which contains all four borders. You can apply different borders to each side of a cell or range. For more information how to apply borders to a range of cells, see Range.Borders property. You can set border properties for an individual border only with Range and Style objects. Other bordered objects, such as error bars and series lines, have a border that's treated as a single entity, regardless of how many sides it has. For these objects, you must return and set properties for the entire border as a unit. For more information, see the Border object.

**Example:**

```vba
Worksheets(1).Range("A1").Borders.LineStyle = xlDouble
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
  Returns or sets a Variant value that represents the color of all four borders.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As Border  (read-only)`  
  Returns a Border object that represents one of the borders of either a range of cells or a style.
- `LineStyle As Variant  (read/write)`  
  Returns or sets the line style for the border. Read/write XlLineStyle, xlGray25, xlGray50, xlGray75, or xlAutomatic.
- `_NewEnum As IUnknown  (read-only)`
- `Value As Variant  (read/write)`  
  Synonym for Borders.LineStyle.
- `Weight As Variant  (read/write)`  
  Returns or sets an XlBorderWeight value that represents the weight of the border.
- `_Default As Border  (read-only)`
- `ThemeColor As Variant  (read/write)`  
  Returns or sets the theme color in the applied color scheme that is associated with the specified object. Read/write Variant.
- `TintAndShade As Variant  (read/write)`  
  Returns or sets a Single that lightens or darkens a color.
