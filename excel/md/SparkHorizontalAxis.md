# SparkHorizontalAxis

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244BB-0000-0000-C000-000000000046}  

Represents the settings for the horizontal axes of a group of sparklines.

**Remarks:** Use the Horizontal property of the SparkAxes object to return the SparkHorizontalAxis object for group of sparklines. The horizontal axis is displayed only if the data for the sparkline has both negative and positive values on the vertical axis.

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent SparklineGroup object for the specified object. Read-only.
- `Axis As SparkColor  (read-only)`  
  Returns a SparkColor object that specifies the color of the horizontal axis of the sparkline. Read-only.
- `IsDateAxis As Boolean  (read-only)`  
  Returns whether the horizontal axis of the sparkline is based on date values. Read-only.
- `RightToLeftPlotOrder As Boolean  (read/write)`  
  Returns or sets whether the points on the horizontal axis are plotted in right-to-left order. Read/write.
