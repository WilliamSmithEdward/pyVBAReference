# SparkVerticalAxis

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244BC-0000-0000-C000-000000000046}  

Represents the settings for the vertical axes of a group of sparklines.

**Remarks:** Use the Vertical property of the SparkAxes object to return the SparkVerticalAxis object for group of sparklines.

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent SparklineGroup object for the specified object. Read-only.
- `MinScaleType As XlSparkScale  (read/write)`  
  Returns or sets how the minimum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group. Read/write.
- `CustomMinScaleValue As Variant  (read/write)`  
  Returns or sets the custom minimum value for the vertical axis of a sparkline. Read/write.
- `MaxScaleType As XlSparkScale  (read/write)`  
  Returns or sets how the maximum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group. Read/write.
- `CustomMaxScaleValue As Variant  (read/write)`  
  Returns or sets the custom maximum value for the vertical axis of a sparkline. Read/write.
