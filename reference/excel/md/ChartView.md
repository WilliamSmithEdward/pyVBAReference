# ChartView

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024488-0000-0000-C000-000000000046}  

Represents a view of a chart.

**Remarks:** The ChartView object is one of the objects that can be returned by the SheetViews collection, similar to the Sheets collection. The ChartView object applies only to chart sheets.

**Example:**

```vba
ActiveWindow.SheetViews.Item(1)
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Sheet As Object  (read-only)`  
  Returns the sheet name for the specified ChartView object. Read-only.
