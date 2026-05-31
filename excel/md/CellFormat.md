# CellFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024450-0000-0000-C000-000000000046}  

Represents the search criteria for the cell format.

**Remarks:** Use the FindFormat or ReplaceFormat properties of the Application object to return a CellFormat object. Use the Borders, Font, or Interior properties of the CellFormat object to define the search criteria for the cell format.

**Example:**

```vba
Sub ChangeCellFormat()

 ' Set the interior of cell A1 to yellow.
 Range("A1").Select
 Selection.Interior.ColorIndex = 36
 MsgBox "The cell format for cell A1 is a yellow interior."

 ' Set the CellFormat object to replace yellow with green.
 With Application
 .FindFormat.Interior.ColorIndex = 36
 .ReplaceFormat.Interior.ColorIndex = 35
 End With

 ' Find and replace cell A1's yellow interior with green.
 ActiveCell.Replace What:="", Replacement:="", LookAt:=xlPart, _
 SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, _
 ReplaceFormat:=True
 MsgBox "The cell format for cell A1 is replaced with a green interior."

End Sub
```

## Properties (18)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Borders As Borders  (read/write)`  
  Returns or sets a Borders collection that represents the search criteria based on the cell's border format.
- `Font As Font  (read/write)`  
  Returns a Font object, allowing the user to set or return the search criteria based on the cell's font format.
- `Interior As Interior  (read/write)`  
  Returns an Interior object allowing the user to set or return the search criteria based on the cell's interior format.
- `NumberFormat As Variant  (read/write)`  
  Returns or sets a Variant value that represents the format code for the object.
- `NumberFormatLocal As Variant  (read/write)`  
  Returns or sets a Variant value that represents the format code for the object as a string in the language of the user.
- `AddIndent As Variant  (read/write)`  
  Returns or sets a Variant value that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically).
- `IndentLevel As Variant  (read/write)`  
  Returns or sets a Variant value that represents the indent level for the cell or range. Can be an integer from 0 to 250.
- `HorizontalAlignment As Variant  (read/write)`  
  Returns or sets a Variant value that represents the horizontal alignment for the specified object.
- `VerticalAlignment As Variant  (read/write)`  
  Returns or sets a Variant value that represents the vertical alignment of the specified object.
- `Orientation As Variant  (read/write)`  
  Returns or sets a Variant value that represents the text orientation.
- `ShrinkToFit As Variant  (read/write)`  
  Returns or sets a Variant value that indicates if text automatically shrinks to fit in the available column width.
- `WrapText As Variant  (read/write)`  
  Returns or sets a Variant value that indicates if Microsoft Excel wraps the text in the object.
- `Locked As Variant  (read/write)`  
  Returns or sets a Variant value that indicates if the object is locked.
- `FormulaHidden As Variant  (read/write)`  
  Returns or sets a Variant value that indicates if the formula will be hidden when the worksheet is protected.
- `MergeCells As Variant  (read/write)`  
  True if the range or style contains merged cells. Read/write Variant.

## Methods (1)

- `Clear()`  
  Clears the criteria set in the FindFormat and ReplaceFormat properties.
