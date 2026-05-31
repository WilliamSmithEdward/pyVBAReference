# SparkPoints

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244B8-0000-0000-C000-000000000046}  

Represents the settings of the markers for points of data on a sparkline.

**Remarks:** Use the SparkPoints object to set the color and visibility of markers for points of data on a sparkline. Use the Points property of the SparklineGroup object to return a SparkPoints object. The properties of the SparkPoints object correspond to the settings of the High Point, Low Point, Negative Point, First Point, Last Point, and Markers check boxes in the Show section, and to the settings of the items on the Marker Color drop-down list in the Style section of the Sparkline Tools Design tab of the ribbon.

## Properties (9)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent SparklineGroup object for the specified object. Read-only.
- `Negative As SparkColor  (read-only)`  
  Returns a SparkColor object that represents the color and visibility of markers for negative points of data on a sparkline. Read-only.
- `Markers As SparkColor  (read-only)`  
  Returns a SparkColor object that represents the color and visibility of markers for positive points of data on a sparkline. Read-only.
- `Highpoint As SparkColor  (read-only)`  
  Returns a SparkColor object that represents the color and visibility of the marker for the highest point of data on a sparkline. Read-only.
- `Lowpoint As SparkColor  (read-only)`  
  Returns a SparkColor object that represents the color and visibility of the marker for the lowest point of data on a sparkline. Read-only.
- `Firstpoint As SparkColor  (read-only)`  
  Returns a SparkColor object that represents the color and visibility of the marker for the first point of data on a sparkline. Read-only.
- `Lastpoint As SparkColor  (read-only)`  
  Returns a SparkColor object that represents the color and visibility of the marker for the last point of data on a sparkline. Read-only.
