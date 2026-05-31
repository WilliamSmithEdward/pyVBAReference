# SparklineGroups

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244B6-0000-0000-C000-000000000046}  

Represents a collection of sparkline groups.

**Remarks:** The SparklineGroups object can contain multiple SparklineGroup objects. Use the SparklineGroups property of the Range object to return an existing SparklineGroups collection from its parent range. Use the Add method to create a group of new sparklines. Use the Group method to create a group of existing sparklines.

**Example:**

```vba
Range("A1:A4").Select
Selection.SparklineGroups.Group Location := Range("A1")
Selection.SparklineGroups.Item(1).Points.Markers.Visible = True
Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 255
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the Range object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns the count of sparkline groups in the associated Range object. Read-only.
- `Item As SparklineGroup  (read-only)`  
  Returns a SparklineGroup object from a collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `_Default As SparklineGroup  (read-only)`

## Methods (5)

- `Add(Type As XlSparkType, SourceData As String) As SparklineGroup`  
  Creates a new sparkline group and returns a SparklineGroup object.
    - `Type As XlSparkType` (required): The type of sparkline.
    - `SourceData As String` (required): Represents the range to use to create the sparkline.
- `Clear()`  
  Clears the selected sparklines.
- `ClearGroups()`  
  Clears the selected sparkline groups.
- `Group(Location As Range)`  
  Groups the selected sparklines.
    - `Location As Range` (required): The location of the first cell in the group.
- `Ungroup()`  
  Ungroups the sparklines in the selected sparkline group.
