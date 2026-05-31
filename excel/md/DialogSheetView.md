# DialogSheetView

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002448A-0000-0000-C000-000000000046}  

Represents the current Dialog sheet view in a workbook.

**Remarks:** To access this object, you must have a dialog sheet that was developed in the active workbook. Without the dialog sheet, the view properties for the object return an empty string value.

**Example:**

```vba
Worksheets("Sheet1").DialogSheetView.Visible = True
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Sheet As Object  (read-only)`  
  Returns the sheet name for the specified DialogSheetView object. Read-only.
