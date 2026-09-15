# ChartGroups

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A5E-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents one or more series plotted in a chart with the same format.

**Remarks:** A ChartGroups collection is a collection of all the ChartGroup objects in the specified chart. A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format. The following example displays the number of chart groups on the first chart of the active document. Use the ChartGroups method to return the ChartGroups collection. The following example adds drop lines to chart group 1 on chart sheet 1. Use ChartGroups (index), where index is the chart group index number, to return a single ChartGroup object.

## Properties (4)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.

## Methods (2)

- `Item(Index As Variant) As ChartGroup`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The index number for the object.
- `_NewEnum() As IUnknown`
