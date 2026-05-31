# Databar

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024496-0000-0000-C000-000000000046}  

Represents a data bar conditional formating rule. Applying a data bar to a range helps you see the value of a cell relative to other cells.

**Remarks:** All conditional formatting objects are contained within a FormatConditions collection object, which is a child of a Range collection. You can create a data bar formatting rule by using either the Add or AddDatabar methods of the FormatConditions collection. You use the MinPoint and MaxPoint properties of the Databar object to set the values of the shortest bar and longest bar of a range of data. These properties return a ConditionValue object, with which you can specify how the thresholds are evaluated. The Databar object also provides properties that enable you to specify an axis line that is displayed when negative values are present, and to specify the color and formatting of data bars.

**Example:**

```vba
Sub CreateDatabarCF()

 Dim cfDatabar As Databar

 ' Create a range of data with a couple of extreme values
 With ActiveSheet
 .Range("D1") = 1
 .Range("D2") = 45
 .Range("D3") = 50
 .Range("D2:D3").AutoFill Destination:=Range("D2:D8")
 .Range("D9") = 500
 End With

 Range("D1:D9").Select

 ' Create a data bar with default behavior
 Set cfDatabar = Selection.FormatConditions.AddDatabar
 MsgBox "Because of the extreme values, middle data bars are very similar"

 ' The MinPoint and MaxPoint properties return a ConditionValue object
 ' which you can use to change threshold parameters
 cfDatabar.MinPoint.Modify newtype:=xlConditionValuePercentile, _
 newvalue:=5
 cfDatabar.MaxPoint.Modify newtype:=xlConditionValuePercentile, _
 newvalue:=75

End Sub
```

## Properties (22)

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
- `MinPoint As ConditionValue  (read-only)`  
  Returns a ConditionValue object that specifies how the shortest bar is evaluated for a data bar conditional format.
- `MaxPoint As ConditionValue  (read-only)`  
  Returns a ConditionValue object that specifies how the longest bar is evaluated for a data bar conditional format.
- `PercentMin As Long  (read/write)`  
  Returns or sets a Long value that specifies the length of the shortest data bar as a percentage of cell width.
- `PercentMax As Long  (read/write)`  
  Returns or sets a Long value that specifies the length of the longest data bar as a percentage of cell width.
- `BarColor As Object  (read-only)`  
  Returns a FormatColor object that you can use to modify the color of the bars in a data bar conditional format.
- `ShowValue As Boolean  (read/write)`  
  Returns or sets a Boolean value that specifies if the value in the cell is displayed if the data bar conditional format is applied to the range.
- `Formula As String  (read/write)`  
  Returns or sets a String representing a formula, which determines the values to which the data bar will be applied.
- `Type As Long  (read-only)`  
  Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Read-only.
- `PTCondition As Boolean  (read-only)`  
  Returns a Boolean value indicating if the conditional format is being applied to a PivotTable. Read-only.
- `ScopeType As XlPivotConditionScope  (read/write)`  
  Returns or sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable.
- `Direction As Long  (read/write)`  
  Returns or sets the direction that the databar is displayed. Read/write.
- `BarFillType As XlDataBarFillType  (read/write)`  
  Returns or sets how a data bar is filled with color. Read/write.
- `AxisPosition As XlDataBarAxisPosition  (read/write)`  
  Returns or sets the position of the axis of the data bars specified by a conditional formatting rule. Read/write.
- `AxisColor As Object  (read-only)`  
  Returns the color of the axis for cells with conditional formatting as data bars. Read-only.
- `BarBorder As DataBarBorder  (read-only)`  
  Returns an object that specifies the border of a data bar. Read-only.
- `NegativeBarFormat As NegativeBarFormat  (read-only)`  
  Returns the NegativeBarFormat object associated with a data bar conditional formatting rule. Read-only.

## Methods (4)

- `SetFirstPriority()`  
  Sets the priority value for this conditional formatting rule to 1 so that it will be evaluated before all other rules on the worksheet.
- `SetLastPriority()`  
  Sets the evaluation order for this conditional formatting rule so that it is evaluated after all other rules on the worksheet.
- `Delete()`  
  Deletes the specified conditional formatting rule object.
- `ModifyAppliesToRange(Range As Range)`  
  Sets the cell range to which this formatting rule applies.
    - `Range As Range` (required): The range to which this formatting rule will be applied.
