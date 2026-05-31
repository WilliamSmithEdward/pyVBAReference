# NegativeBarFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244BF-0000-0000-C000-000000000046}  

Represents the color settings of the data bars for negative values that are defined by a data bar conditional formating rule.

**Remarks:** The properties of the NegativeBarFormat object can be used to specify the fill color and border of the data bars for negative values. Use the NegativeBarFormat property of the DataBar object that represents a data bar conditional formatting rule to access the NegativeBarFormat object associated with that rule.

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ColorType As XlDataBarNegativeColorType  (read/write)`  
  Specifies whether to use the same fill color as positive data bars. Read/write.
- `BorderColorType As XlDataBarNegativeColorType  (read/write)`  
  Specifies whether to use the same border color as positive data bars. Read/write.
- `Color As Object  (read-only)`  
  Returns a FormatColor object that you can use to specify the fill color for negative data bars. Read-only.
- `BorderColor As Object  (read-only)`  
  Returns a FormatColor object that you can use to specify the border color for negative data bars. Read-only.
