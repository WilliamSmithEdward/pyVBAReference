# SparkColor

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244BD-0000-0000-C000-000000000046}  

Represents the color of the horizontal axis and the markers for points in a sparkline.

**Remarks:** The SparkColor object corresponds to the settings of the items available on the Marker Color drop-down list in the Style section of the Sparkline Tools Design tab of the ribbon. Use the corresponding properties of the SparkPoints object to set the colors of these items.

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent SparklineGroup object for the specified object. Read-only.
- `Visible As Boolean  (read/write)`  
  Returns or sets whether the point is visible. Read/write.
- `Color As FormatColor  (read-only)`  
  Returns a FormatColor object that you can use to set the color of the horizontal axis or the markers for points in a sparkline. Read-only.
