# Sparkline

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244B9-0000-0000-C000-000000000046}  

Represents a single sparkline.

**Remarks:** Use the ModifyLocation method to change the location of a single sparkline, and use the ModifySourceData method to change the range of the source data. To work with a group of sparklines at the same time, use the members of the SparklineGroup object.

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent SparklineGroup object for the specified object. Read-only.
- `Location As Range  (read/write)`  
  Returns or sets the location of a single sparkline. Read/write.
- `SourceData As String  (read/write)`  
  Returns or sets the range that contains the source data for a single sparkline. Read/write.

## Methods (2)

- `ModifyLocation(Range As Range)`  
  Modifies the location of a single sparkline.
    - `Range As Range` (required): The cell that contains the sparkline.
- `ModifySourceData(Formula As String)`  
  Modifies the source data for a single sparkline.
    - `Formula As String` (required): The range that contains the source data.
