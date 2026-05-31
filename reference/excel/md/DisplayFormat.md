# DisplayFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244C2-0000-0000-C000-000000000046}  

Represents the display settings for an associated Range object. Read-only.

**Remarks:** Actions such as changing the conditional formatting or table style of a range can cause what is displayed in the current user interface to be inconsistent with the values in the corresponding properties of the Range object. Use the properties of the DisplayFormat object to return the values as they are displayed in the current user interface.

## Properties (21)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Borders As Borders  (read-only)`  
  Returns a Borders object that represents the borders of the associated Range object as it is displayed in the current user interface. Read-only.
- `Characters As Characters  (read-only)`  
  Returns a Characters object that represents a range of characters within the text of the associated Range object as it is displayed in the current user interface. Read-only.
- `Font As Font  (read-only)`  
  Returns a Font object that represents the font of the associated Range as it is displayed in the current user interface. Read-only.
- `Style As Variant  (read-only)`  
  Returns a value, containing a Style object, that represents the style of the associated Range object as it is displayed in the current user interface.
- `AddIndent As Variant  (read-only)`  
  Returns a value that indicates if Microsoft Excel automatically indents text of the associated Range object when the text alignment in a cell is set to equal distribution (either horizontally or vertically), as it is displayed in the current user interface. Read-only.
- `FormulaHidden As Variant  (read-only)`  
  Returns a value that indicates if the formula of the associated Range object is hidden when the worksheet is protected as it is displayed in the current user interface. Read-only.
- `HorizontalAlignment As Variant  (read-only)`  
  Returns a value that represents the horizontal alignment of the associated Range object as it is displayed in the current user interface. Read-only.
- `IndentLevel As Variant  (read-only)`  
  Returns a value that represents the indent level of the associated Range object as it is displayed in the current user interface. Read-only.
- `Interior As Interior  (read-only)`  
  Returns an Interior object that represents the interior of the associated Range object as it is displayed in the current user interface. Read-only.
- `Locked As Variant  (read-only)`  
  Returns a value that indicates if the associated Range object is locked as it is displayed in the current user interface. Read-only.
- `MergeCells As Variant  (read-only)`  
  Returns a value that indicates if the associated Range object contains merged cells as it is displayed in the current user interface. Read-only.
- `NumberFormat As Variant  (read-only)`  
  Returns a value that represents the format code of the associated Range object as it is displayed in the current user interface. Read-only.
- `NumberFormatLocal As Variant  (read-only)`  
  Returns a value that represents the format code of the associated Range object as a string in the language of the user as it is displayed in the current user interface. Read-only.
- `Orientation As Variant  (read-only)`  
  Returns a value that represents the text orientation of the associated Range object as it is displayed in the current user interface. Read-only.
- `ReadingOrder As Long  (read-only)`  
  Returns the reading order of the associated Range object as it is displayed in the current user interface. Read-only.
- `ShrinkToFit As Variant  (read-only)`  
  Returns a value that indicates if Microsoft Excel automatically shrinks text to fit in the available column width of the associated Range object as it is displayed in the current user interface. Read-only.
- `VerticalAlignment As Variant  (read-only)`  
  Returns a value that represents the vertical alignment of the associated Range object as it is displayed in the current user interface. Read-only.
- `WrapText As Variant  (read-only)`  
  Returns a value that indicates if Microsoft Excel wraps the text of the associated Range object as it is displayed in the current user interface. Read-only.
