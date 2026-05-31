# ColorScale

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024493-0000-0000-C000-000000000046}  

Represents a color scale conditional formatting rule.

**Remarks:** All conditional formatting objects are contained within a FormatConditions collection object, which is a child of a Range collection. You can create a color scale formatting rule by using either the Add or AddColorScale method of the FormatConditions collection. Color scales are visual guides that help you understand data distribution and variation. You can apply either a two-color or a three-color scale to a range of data, data in a table, or data in a PivotTable report. For a two-color scale conditional format, you assign the value, type, and color to the minimum and maximum thresholds of a range. A three-color scale also has a midpoint threshold. Each of these thresholds is determined by setting the properties of the ColorScaleCriteria object. The ColorScaleCriteria object, which is a child of the ColorScale object, is a collection of all of the ColorScaleCriterion objects for the color scale.

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

## Properties (11)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Priority As Long  (read/write)`  
  Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist on a worksheet.
- `StopIfTrue As Boolean  (read-only)`  
  Returns or sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True.
- `AppliesTo As Range  (read-only)`  
  Returns a Range object specifying the cell range to which the formatting rule is applied.
- `Formula As String  (read/write)`  
  Returns or sets a String representing a formula that determines the values to which the icon set will be applied.
- `Type As Long  (read-only)`  
  Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Read-only.
- `PTCondition As Boolean  (read-only)`  
  Returns a Boolean value indicating if the conditional format is being applied to a PivotTable. Read-only.
- `ScopeType As XlPivotConditionScope  (read/write)`  
  Returns or sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable.
- `ColorScaleCriteria As ColorScaleCriteria  (read-only)`  
  Returns a ColorScaleCriteria object, which is a collection of individual ColorScaleCriterion objects. The ColorScaleCriterion object specifies the type, value, and color of threshold criteria used in the color scale conditional format. Read-only.

## Methods (4)

- `SetFirstPriority()`  
  Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.
- `SetLastPriority()`  
  Sets the evaluation order for this conditional formatting rule so that it is evaluated after all other rules on the worksheet.
- `Delete()`  
  Deletes the specified conditional formatting rule object.
- `ModifyAppliesToRange(Range As Range)`  
  Sets the cell range to which this formatting rule applies.
    - `Range As Range` (required): The range to which this formatting rule will be applied.
