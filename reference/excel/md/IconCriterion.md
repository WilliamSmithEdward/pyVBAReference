# IconCriterion

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024499-0000-0000-C000-000000000046}  

Represents the criterion for an individual icon in an icon set. The criterion specifies the range of values and the threshold type associated with the icon in an icon set conditional formatting rule.

**Remarks:** All the criteria for an icon set conditional format are contained in an IconCriteria collection. You can access each IconCriterion object in the collection by passing an index into the collection. See the example for details.

**Example:**

```vba
Sub CreateIconSetCF()

    Dim cfIconSet As IconSetCondition

    'Fill cells with sample data
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

## Properties (5)

- `Index As Long  (read-only)`  
  Returns a Long value indicating which threshold the criteria represents. Read-only.
- `Type As XlConditionValueTypes  (read/write)`  
  Returns one of the constants of the XlConditionValueTypes enumeration, which specifies how the threshold value for an icon set is determined. Read-only.
- `Value As Variant  (read/write)`  
  Returns or sets the threshold value for an icon in a conditional format. Read/write Variant.
- `Operator As Long  (read/write)`  
  Returns or sets one of the constants of the XlFormatConditionOperator enumeration, which specifies if the threshold is "greater than" or "greater than or equal to" the threshold value.
- `Icon As XlIcon  (read/write)`  
  Returns or specifies the icon for a criterion in an icon set conditional formatting rule. Read/write.
