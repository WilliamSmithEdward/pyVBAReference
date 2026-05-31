# IconSetCondition

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024497-0000-0000-C000-000000000046}  

Represents an icon set conditional formatting rule.

**Remarks:** All conditional formatting objects are contained within a FormatConditions collection object, which is a child of a Range collection. You can create an icon set formatting rule by using either the Add method or AddIconSetCondition method of the FormatConditions collection. Each icon set contains three, four, or five icons. You use the IconSets property of the Workbook object to return an IconSets object to specify one of the built-in icon sets. Each individual icon in the icon set is then assigned to a subset of the values of the range by the members of the IconCriteria object. The type of threshold is also specified by this object.

**Example:**

```vba
Sub CreateIconSetCF()

    Dim cfIconSet As IconSetCondition

    'Fill cells with sample data from 1 to 10
    With ActiveSheet
        .Range("C1") = 55
        .Range("C2") = 92
        .Range("C3") = 88
        .Range("C4") = 77
        .Range("C5") = 66
        .Range("C6") = 93
        .Range("C7") = 76
        .Range("C8") = 80
        .Range("C9") = 79
        .Range("C10") = 83
        .Range("C11") = 66
        .Range("C12") = 74
    End With

    Range("C1:C12").Select

    'Create an icon set conditional format for the created sample data range
    Set cfIconSet = Selection.FormatConditions.AddIconSetCondition

    'Change the icon set to a five-arrow icon set
    cfIconSet.IconSet = ActiveWorkbook.IconSets(xl5Arrows)

    'The IconCriterion collection contains all IconCriteria
    'By indexing into the collection you can modify each criterion

    With cfIconSet.IconCriteria(1)
        .Type = xlConditionValueNumber
        .Value = 0
        .Operator = 7
    End With
    With cfIconSet.IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 60
        .Operator = 7
    End With
    With cfIconSet.IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 70
        .Operator = 7
    End With
    With cfIconSet.IconCriteria(4)
        .Type = xlConditionValueNumber
        .Value = 80
        .Operator = 7
    End With
    With cfIconSet.IconCriteria(5)
        .Type = xlConditionValueNumber
        .Value = 90
        .Operator = 7
    End With

End Sub
```

## Properties (15)

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
- `Type As Long  (read-only)`  
  Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Read-only.
- `PTCondition As Boolean  (read-only)`  
  Returns a Boolean value indicating if the conditional format is being applied to a PivotTable. Read-only.
- `ScopeType As XlPivotConditionScope  (read/write)`  
  Returns or sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable.
- `ReverseOrder As Boolean  (read/write)`  
  Returns or sets a Boolean value indicating if the order of icons is reversed for an icon set.
- `PercentileValues As Boolean  (read/write)`  
  Returns or sets a Boolean value indicating if the thresholds for an icon set conditional format are determined by using percentiles.
- `ShowIconOnly As Boolean  (read/write)`  
  Returns or sets a Boolean value indicating if only the icon is displayed for an icon set conditional format.
- `Formula As String  (read/write)`  
  Returns or sets a String representing a formula, which determines the values to which the icon set will be applied.
- `IconSet As Variant  (read/write)`  
  Returns or sets an IconSets collection, which specifies the icon set used in the conditional format.
- `IconCriteria As IconCriteria  (read-only)`  
  Returns an IconCriteria collection, which represents the set of criteria for an icon set conditional formatting rule.

## Methods (4)

- `ModifyAppliesToRange(Range As Range)`  
  Sets the cell range to which this formatting rule applies.
    - `Range As Range` (required): The range to which this formatting rule will be applied.
- `SetFirstPriority()`  
  Sets the priority value for this conditional formatting rule to 1 so that it will be evaluated before all other rules on the worksheet.
- `SetLastPriority()`  
  Sets the evaluation order for this conditional formatting rule so that it is evaluated after all other rules on the worksheet.
- `Delete()`  
  Deletes the specified conditional formatting rule object.
