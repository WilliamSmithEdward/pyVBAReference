# PivotLine

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024480-0000-0000-C000-000000000046}  

A PivotLine object is a line of rows or columns in an Excel PivotTable.

**Remarks:** PivotLines contain only visible items, so collapsed children of items and items in hidden levels are not present in the PivotLines collection. PivotLines always have a PivotItem in all positions. This means that the PivotLines representing subtotals in the PivotTable contain fewer PivotItems than regular PivotLines.

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified PivotLine object. Read-only.
- `LineType As XlPivotLineType  (read-only)`  
  Returns an XlPivotLineType constant that indicates the type of PivotLine. Read-only.
- `Position As Long  (read-only)`  
  Returns or sets the position of the PivotLine object. Read-only.
- `PivotLineCells As PivotLineCells  (read-only)`  
  Returns a collection of PivotCell objects in a PivotLine. Read-only.
- `PivotLineCellsFull As PivotLineCells  (read-only)`  
  Retrieves all the PivotLine cells including those that are hidden in compact form. Read-only.
