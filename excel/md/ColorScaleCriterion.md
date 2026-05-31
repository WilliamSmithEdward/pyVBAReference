# ColorScaleCriterion

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024495-0000-0000-C000-000000000046}  

Represents the criteria for the minimum, midpoint, or maximum thresholds for a color format conditional format.

**Remarks:** Each ColorScaleCriterion is part of the ColorScaleCriteria collection. Use ColorScaleCriteria (_index_), to return an individual criterion. You can set the Type, Value, and FormatColor of each threshold for the color scale.

**Example:**

```vba
Sub CreateColorScaleCF()

 Dim cfColorScale As ColorScale

 'Fill cells with sample data from 1 to 10
 With ActiveSheet
 .Range("C1") = 1
 .Range("C2") = 2
 .Range("C1:C2").AutoFill Destination:=Range("C1:C10")
 End With

 Range("C1:C10").Select

 'Create a two-color ColorScale object for the created sample data range
 Set cfColorScale = Selection.FormatConditions.AddColorScale(ColorScaleType:=2)

 'Set the minimum threshold to red and maximum threshold to blue
 cfColorScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0)
 cfColorScale.ColorScaleCriteria(2).FormatColor.Color = RGB(0, 0, 255)

End Sub
```

## Properties (4)

- `Index As Long  (read-only)`  
  Returns a Long value indicating which threshold the criteria represents. Read-only.
- `Type As XlConditionValueTypes  (read/write)`  
  Returns one of the constants of the XlConditionValueTypes enumeration, which specifies how the threshold values for a data bar or color scale conditional format are determined. Read-only.
- `Value As Variant  (read/write)`  
  Returns or sets the minimum, midpoint, or maximum threshold value for a color scale conditional format. Read/write Variant.
- `FormatColor As FormatColor  (read-only)`  
  Returns a FormatColor object, which specifies the color assigned to the threshold of a color scale conditional format. Read-only.
