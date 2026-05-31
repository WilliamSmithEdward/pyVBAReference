# TreeviewControl

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002444B-0000-0000-C000-000000000046}  

Represents the hierarchical member-selection control of a cube field.

**Remarks:** Use this object primarily for macro recording; it is not intended for any other use.

**Example:**

```vba
ActiveSheet.PivotTables("PivotTable2") _
 .CubeFields(1).TreeviewControl.Drilled = _
 Array(Array("", ""), _
 Array("[state].[states].[CA]", _
 "[state].[states].[MD]"))
```

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Hidden As Variant  (read/write)`  
  Returns or sets a Variant value that represents the hidden status of the cube field members in the hierarchical member-selection control of a cube field.
- `Drilled As Variant  (read/write)`  
  Sets the "drilled" (expanded or visible) status of the cube field members in the hierarchical member-selection control of a cube field. This property is used primarily for macro recording and isn't intended for any other use. Read/write.
