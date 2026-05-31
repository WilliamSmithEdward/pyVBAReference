# ChartColorFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024436-0000-0000-C000-000000000046}  

Represents a foreground or background color.

**Remarks:** Use the ForeColor property to return a ChartColorFormat object that represents the foreground fill color. Use the BackColor property to return the background fill color. Use the RGB property to return the color as an explicit red-green-blue value. Use the SchemeColor property to return or set the color as one of the colors in the current color scheme.

**Example:**

```vba
With myChart.ChartArea.Fill
    .Visible = True
    .ForeColor.SchemeColor = 15
    .BackColor.SchemeColor = 17
    .TwoColorGradient msoGradientHorizontal, 1
End With
```

## Properties (7)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `SchemeColor As Long  (read/write)`
- `RGB As Long  (read-only)`
- `_Default As Long  (read-only)`
- `Type As Long  (read-only)`
