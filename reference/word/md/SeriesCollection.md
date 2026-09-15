# SeriesCollection

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {8FEB78F7-35C6-4871-918C-193C3CDD886D}  

Represents a collection of all the Series objects in the specified chart or chart group.

**Remarks:** Use the SeriesCollection method to return the SeriesCollection collection.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.SeriesCollection.Extend "='Sheet1'!$C$6:$C$10"
 End If
End With
```

## Properties (4)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (6)

- `Add(Source As Variant, [Rowcol As XlRowCol], [SeriesLabels As Variant], [CategoryLabels As Variant], [Replace As Variant]) As Series`  
  Adds one or more new series to the collection.
    - `Source As Variant` (required): The new data as a string representation of a range contained in the Workbook property of the ChartData object for the chart.
    - `Rowcol As XlRowCol` (optional): One of the enumeration values that specifies whether the new values are in the rows or columns of the specified range.
    - `SeriesLabels As Variant` (optional): True if the first row or column contains the name of the data series. False if the first row or column contains the first data point of the series. If this argument is omitted, Microsoft Word attempts to determine the location of the series name from the contents of the first row or column.
    - `CategoryLabels As Variant` (optional): True if the first row or column contains the name of the category labels. False if the first row or column contains the first data point of the series. If this argument is omitted, Word attempts to determine the location of the category label from the contents of the first row or column.
    - `Replace As Variant` (optional): If CategoryLabels is True and Replace is True, the specified categories replace the categories that currently exist for the series. If Replace is False, the existing categories will not be replaced. The default is False.
- `Extend(Source As Variant, [Rowcol As Variant], [CategoryLabels As Variant]) As Variant`  
  Adds new data points to an existing series collection.
    - `Source As Variant` (required): The new data to be added to the SeriesCollection object, represented as an A1-style range reference.
    - `Rowcol As Variant` (optional): One of the XlRowCol enumeration values that specifies whether the new values are in the rows or columns of the given range source. If this argument is omitted, Microsoft Word attempts to determine where the values are by the size and orientation of the selected range or by the dimensions of the array.
    - `CategoryLabels As Variant` (optional): True to have the first row or column contain the name of the category labels. False to have the first row or column contain the first data point of the series. If this argument is omitted, Word attempts to determine the location of the category label from the contents of the first row or column.
- `Item(Index As Variant) As Series`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
- `NewSeries() As Series`  
  Creates a new series.
- `_Default(Index As Variant) As Series`
