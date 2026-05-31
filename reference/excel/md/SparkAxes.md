# SparkAxes

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244BA-0000-0000-C000-000000000046}  

Represents the settings for the horizontal and vertical axes of a group of sparklines.

**Remarks:** Use the Axes property of a SparklineGroup object to return the SparkAxes object for that group of sparklines.

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent SparklineGroup object for the specified SparkAxes object. Read-only.
- `Vertical As SparkVerticalAxis  (read-only)`  
  Returns the SparkVerticalAxis object for the specified SparkAxes object. Read-only.
- `Horizontal As SparkHorizontalAxis  (read-only)`  
  Returns the SparkHorizontalAxis object for the specified SparkAxes object. Read-only.
