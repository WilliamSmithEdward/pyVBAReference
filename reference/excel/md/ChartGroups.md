# ChartGroups

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002085A-0000-0000-C000-000000000046}  

Represents one or more series plotted in a chart with the same format.

**Remarks:** A ChartGroups collection is a collection of all the ChartGroup objects in the specified chart. A chart contains one or more chart groups, each chart group contains one or more Series objects, and each series contains one or more Points objects. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format. Use the ChartGroups method of the Chart object to return the ChartGroups collection. The following example displays the number of chart groups on embedded chart 1 on worksheet 1. Use ChartGroups (_index_), where _index_ is the chart-group index number, to return a single ChartGroup object. The following example adds drop lines to chart group 1 on chart sheet 1. If the chart has been activated, you can use ActiveChart. Because the index number for a particular chart group can change if the chart format used for that group is changed, it may be easier to use one of the named chart group shortcut methods to return a particular chart group.

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (2)

- `Item(Index As Variant) As ChartGroup`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The index number for the object.
- `_NewEnum() As IUnknown`
