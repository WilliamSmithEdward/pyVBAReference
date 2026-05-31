# IconSet

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002449B-0000-0000-C000-000000000046}  

Represents a single set of icons that are used in an icon set conditional formatting rule.

**Remarks:** The IconSet object is a child object of the IconSets collection. The icon set for the conditional format is assigned by using the IconSet property of the IconSetCondition object. You set this property to one of the built-in icon sets by passing one of the constants of the XlIconSet enumeration as an index of the IconSets property of the Workbook object. See the example for details.

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

 'Change the icon set to a 5-arrow icon set
 cfIconSet.IconSet = ActiveWorkbook.IconSets(xl5Arrows)

 'The IconCriterion collection contains all of IconCriteria
 'By indexing into the collection you can modify each criteria

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

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ID As XlIconSet  (read-only)`  
  Returns one of the constants of the XlIconSet enumeration, which specifies the name of the icon set used in an icon set conditional formatting rule.
- `Count As Long  (read-only)`  
  Returns a Long value that specifies the number of icons in the icon set. Read-only.
- `_Default As Icon  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Item As Icon  (read-only)`  
  Returns an Icon object that represents a single icon from an icon set. Read-only.
