# ConditionValue

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024492-0000-0000-C000-000000000046}  

Represents how the shortest bar or longest bar is evaluated for a data bar conditional formatting rule.

**Remarks:** The ConditionValue object is returned by using either the MaxPoint or MinPoint property of the DataBar object. You can change the type of evaluation from the default setting (lowest value for the shortest bar and highest value for the longest bar) by using the Modify method.

**Example:**

```vba
Sub CreateDataBarCF()

 Dim cfDataBar As DataBar

 'Create a range of data with a couple of extreme values
 With ActiveSheet
 .Range("D1") = 1
 .Range("D2") = 45
 .Range("D3") = 50
 .Range("D2:D3").AutoFill Destination:=Range("D2:D8")
 .Range("D9") = 500
 End With

 Range("D1:D9").Select

 'Create a data bar with default behavior
 Set cfDataBar = Selection.FormatConditions.AddDatabar
 MsgBox "Because of the extreme values, middle data bars are very similar"

 'The MinPoint and MaxPoint properties return a ConditionValue object
 'which you can use to change threshold parameters
 cfDataBar.MinPoint.Modify newtype:=xlConditionValuePercentile, _
 newvalue:=5
 cfDataBar.MaxPoint.Modify newtype:=xlConditionValuePercentile, _
 newvalue:=75

End Sub
```

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Type As XlConditionValueTypes  (read-only)`  
  Returns one of the constants of the XlConditionValueTypes enumeration, which specifies how the threshold values for a data bar, color scale, or icon set conditional format are determined. Read-only.
- `Value As Variant  (read-only)`  
  Returns or sets the shortest bar or longest bar threshold value for a data bar conditional format. Read/write Variant.

## Methods (1)

- `Modify(newtype As XlConditionValueTypes, [newvalue As Variant])`  
  Modifies how the longest bar or shortest bar is evaluated for a data bar conditional formatting rule.
    - `newtype As XlConditionValueTypes` (required): Specifies how the shortest bar or longest bar is evaluated. The default value is xlConditionLowestValue for the shortest bar and xlConditionHighestValue for the longest bar.
    - `newvalue As Variant` (optional): The value assigned to the shortest or longest data bar. Depending on the _NewType_ argument, this can be a number or a formula that evaluates to a number.
