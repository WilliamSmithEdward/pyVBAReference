# DataBarBorder

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244BE-0000-0000-C000-000000000046}  

Represents the border of the data bars specified by a conditional formatting rule.

**Remarks:** Use the DataBarBorder object to get or set the color and border type for data bars. To access the DataBarBorder object associated with a data bar conditional formatting rule, use the BarBorder property. After retrieving the DataBarBorder object, use its Color property to return a FormatColor object that you can use to set the color of the data bars.

**Example:**

```vba
Range("A1:A10").Select
Range("A1:A10").Activate

Set myDataBar = Selection.FormatConditions.AddDatabar
With myDataBar.BarBorder
 .Type = xlDataBarBorderSolid
 .Color.ThemeColor = xlThemeColorAccent2
 .Color.TintAndShade = 0
End With
```

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Type As XlDataBarBorderType  (read/write)`  
  Returns or sets the type of border for data bars specified by a conditional formatting rule. Read/write.
- `Color As Object  (read-only)`  
  Returns an object that specifies the color of the border of data bars specified by a conditional formatting rule. Read-only.
