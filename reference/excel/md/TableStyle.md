# TableStyle

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244A7-0000-0000-C000-000000000046}  

Represents a single style that can be applied to a table or slicer.

**Remarks:** A table style defines the formatting for one or all of the elements of a table, PivotTable, or slicer. For example, a column is an element of a table. A table style can define that columns in a table are formatted with alternating formatting (also known as banding or stripes).

## Properties (12)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As String  (read-only)`
- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `NameLocal As String  (read-only)`  
  Returns or sets the name of the object in the language of the user. Read-only String.
- `BuiltIn As Boolean  (read-only)`  
  True if the style is a built-in style. Read-only Boolean.
- `TableStyleElements As TableStyleElements  (read-only)`  
  Returns the TableStyleElements object. Read-only.
- `ShowAsAvailableTableStyle As Boolean  (read/write)`  
  Returns or sets a table style shown as available in the table styles gallery. Read/write Boolean.
- `ShowAsAvailablePivotTableStyle As Boolean  (read/write)`  
  Sets or returns whether a style is shown in the gallery for PivotTable styles. Read/write Boolean.
- `ShowAsAvailableSlicerStyle As Boolean  (read/write)`  
  Returns or sets whether the specified table style is shown as available in the slicer styles gallery. Read/write.
- `ShowAsAvailableTimelineStyle As Boolean  (read/write)`  
  Returns or sets whether the specified table style is shown as available in the timeline styles gallery. Read/write Boolean.

## Methods (2)

- `Delete()`  
  Deletes the TableStyle object.
- `Duplicate([NewTableStyleName As Variant]) As TableStyle`  
  Duplicates the TableStyle object and returns a reference to the new copy.
    - `NewTableStyleName As Variant` (optional): The name of the new table style.
