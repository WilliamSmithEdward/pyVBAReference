# WorksheetView

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024487-0000-0000-C000-000000000046}  

An object that defines the behavior of a single worksheet view.

**Remarks:** Controls the look and feel of the application or workbook-level view by providing properties such as DisplayFormulas, DisplayGridlines, and DisplayHeadings.

## Properties (10)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Sheet As Object  (read-only)`  
  Returns the sheet name for the specified WorksheetView object. Read-only.
- `DisplayGridlines As Boolean  (read/write)`  
  True if gridlines are displayed. Read/write Boolean.
- `DisplayFormulas As Boolean  (read/write)`  
  Returns or sets if formulas should be displayed or hidden in the current worksheet view. Read/write Boolean.
- `DisplayHeadings As Boolean  (read/write)`  
  True if both row and column headings are displayed; False if no headings are displayed. Read/write Boolean.
- `DisplayOutline As Boolean  (read/write)`  
  True if outline symbols are displayed. Read/write Boolean.
- `DisplayZeros As Boolean  (read/write)`  
  True if zero values are displayed. Read/write Boolean.
- `DisplayDataTypeIcons As Boolean  (read/write)`
