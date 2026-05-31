# Style

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020852-0000-0000-C000-000000000046}  

Represents a style description for a range.

**Remarks:** The Style object contains all style attributes (font, number format, alignment, and so on) as properties. There are several built-in styles, including Normal, Currency, and Percent. Using the Style object is a fast and efficient way to change several cell-formatting properties on multiple cells at the same time. For the Workbook object, the Style object is a member of the Styles collection. The Styles collection contains all the defined styles for the workbook. You can change the appearance of a cell by changing properties of the style applied to that cell. Keep in mind, however, that changing a style property affects all cells already formatted with that style. Styles are sorted alphabetically by style name. The style index number denotes the position of the specified style in the sorted list of style names. Styles(1) is the first style in the alphabetic list, and Styles(Styles.Count) is the last one in the list. For more information about creating and modifying a style, see the Styles object.

**Example:**

```vba
Worksheets("Sheet1").Range("A1:A10").Style = "Percent"
```

## Properties (30)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `AddIndent As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically).
- `BuiltIn As Boolean  (read-only)`  
  True if the style is a built-in style. Read-only Boolean.
- `Borders As Borders  (read-only)`  
  Returns a Borders collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).
- `Font As Font  (read-only)`  
  Returns a Font object that represents the font of the specified object.
- `FormulaHidden As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if the formula will be hidden when the worksheet is protected.
- `HorizontalAlignment As XlHAlign  (read/write)`  
  Returns or sets an XlHAlign value that represents the horizontal alignment for the specified object.
- `IncludeAlignment As Boolean  (read/write)`  
  True if the style includes the AddIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and Orientation properties of the Style object. Read/write Boolean.
- `IncludeBorder As Boolean  (read/write)`  
  True if the style includes the Color, ColorIndex, LineStyle, and Weight properties of the Border object. Read/write Boolean.
- `IncludeFont As Boolean  (read/write)`  
  True if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties. Read/write Boolean.
- `IncludeNumber As Boolean  (read/write)`  
  True if the style includes the NumberFormat property. Read/write Boolean.
- `IncludePatterns As Boolean  (read/write)`  
  True if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex properties of the Interior object. Read/write Boolean.
- `IncludeProtection As Boolean  (read/write)`  
  True if the style includes the FormulaHidden and Locked protection properties. Read/write Boolean.
- `IndentLevel As Long  (read/write)`  
  Returns or sets a Long value that represents the indent level for the style.
- `Interior As Interior  (read-only)`  
  Returns an Interior object that represents the interior of the specified object.
- `Locked As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if the object is locked.
- `MergeCells As Variant  (read/write)`  
  True if the style contains merged cells. Read/write Variant.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `NameLocal As String  (read-only)`  
  Returns or sets the name of the object, in the language of the user. Read-only String.
- `NumberFormat As String  (read/write)`  
  Returns or sets a String value that represents the format code for the object.
- `NumberFormatLocal As String  (read/write)`  
  Returns or sets a String value that represents the format code for the object as a string in the language of the user.
- `Orientation As XlOrientation  (read/write)`  
  Returns or sets an XlOrientation value that represents the text orientation.
- `ShrinkToFit As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if text automatically shrinks to fit in the available column width.
- `Value As String  (read-only)`  
  Returns a String value that represents the name of the specified style.
- `VerticalAlignment As XlVAlign  (read/write)`  
  Returns or sets an XlVAlign value that represents the vertical alignment of the specified object.
- `WrapText As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if Microsoft Excel wraps the text in the object.
- `_Default As String  (read-only)`
- `ReadingOrder As Long  (read/write)`  
  Returns or sets the reading order for the specified object. Can be one of the following XlReadingOrder constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext. Read/write Long.

## Methods (1)

- `Delete() As Variant`  
  Deletes the object.
