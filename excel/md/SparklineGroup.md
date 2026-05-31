# SparklineGroup

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244B7-0000-0000-C000-000000000046}  

Represents a group of sparklines.

**Remarks:** The SparklineGroup object can contain multiple sparklines, and contains the property settings for the group, such as color and axis settings. Each sparkline is represented by a Sparkline object. Use the Modify method to add or remove sparklines from the sparkline group. Use the ModifyLocation method to change the location of the sparkline, and use the ModifySourceData method to change the range of the source data.

**Example:**

```vba
Dim mySG As SparklineGroup
Set mySG = Range("$A$1:$A$4").SparklineGroups.Add(Type:=xlSparkColumn, SourceData:= _
 "Sheet2!B1:E4")

mySG.SeriesColor.Color = RGB(255, 0, 0)
```

## Properties (17)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of sparklines in the sparkline group. Read-only.
- `Item As Sparkline  (read-only)`  
  Returns a Sparkline object. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Location As Range  (read/write)`  
  Gets or sets the Range object that represents the location of the sparkline group. Read/write.
- `SourceData As String  (read/write)`  
  Returns or sets the range that contains the source data for the sparkline group. Read/write.
- `DateRange As String  (read/write)`  
  Gets or sets the date range for the sparkline group. Read/write.
- `Type As XlSparkType  (read/write)`  
  Gets or sets the type of sparkline for the group. Read/write.
- `SeriesColor As FormatColor  (read-only)`  
  Returns a FormatColor object that represents the main series color for the sparkline group. Read-only.
- `Points As SparkPoints  (read-only)`  
  Returns the associated SparkPoints object for the sparkline group. Read-only.
- `Axes As SparkAxes  (read-only)`  
  Returns the associated SparkAxes object. Read-only.
- `DisplayBlanksAs As XlDisplayBlanksAs  (read/write)`
- `DisplayHidden As Boolean  (read/write)`  
  Specifies if hidden cells are plotted in the sparkline group. Read/write.
- `LineWeight As Variant  (read/write)`  
  Gets or sets the thickness of the sparklines in the sparkline group. Read/write.
- `PlotBy As XlSparklineRowCol  (read/write)`  
  Returns or sets how to plot the sparkline when the data on which it is based is in a square-shaped range. Read/write.

## Methods (5)

- `ModifyLocation(Location As Range)`  
  Sets the associated Range object to modify the location of the sparkline group.
    - `Location As Range` (required): The range that represents the location of the sparkline group.
- `ModifySourceData(SourceData As String)`  
  Sets the range that represents the source data for the sparkline group.
    - `SourceData As String` (required): The range that represents the source data.
- `Modify(Location As Range, SourceData As String)`  
  Sets the location and the source data for the sparkline group.
    - `Location As Range` (required): The Range object that represents the location of the sparkline group.
    - `SourceData As String` (required): The range that represents the source data for the sparkline group.
- `ModifyDateRange(DateRange As String)`  
  Sets the date range for the sparkline group.
    - `DateRange As String` (required): The date range for the sparkline group.
- `Delete()`  
  Deletes the sparkline group.
