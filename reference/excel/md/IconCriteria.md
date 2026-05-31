# IconCriteria

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024498-0000-0000-C000-000000000046}  

Represents the collection of IconCriterion objects. Each IconCriterion object represents the values and threshold type for each icon in an icon set conditional formatting rule.

**Remarks:** The IconCriteria collection is returned from the IconCriteria property of the IconSetCondition object. You can access each IconCriterion object in the collection by passing an index into the collection. See the example for details.

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

## Properties (4)

- `Count As Long  (read-only)`  
  Returns a Long value that specifies the number of criteria for an icon set conditional formatting rule. Read-only.
- `_Default As IconCriterion  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Item As IconCriterion  (read-only)`  
  Returns a single IconCriterion object from the IconCriteria collection. Read-only.
