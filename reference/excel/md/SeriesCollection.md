# SeriesCollection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002086C-0000-0000-C000-000000000046}  

A collection of all the Series objects in the specified chart or chart group.

**Remarks:** Use the SeriesCollection method of the Chart property to return the SeriesCollection collection.

**Example:**

```vba
Worksheets(1).ChartObjects(1).Chart. _
 SeriesCollection.Extend Worksheets(1).Range("c1:c10")
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (7)

- `Add(Source As Variant, [Rowcol As XlRowCol], [SeriesLabels As Variant], [CategoryLabels As Variant], [Replace As Variant]) As Series`  
  Adds one or more new series to the SeriesCollection collection.
    - `Source As Variant` (required): The new data as a Range object.
    - `Rowcol As XlRowCol` (optional): Specifies whether the new values are in the rows or columns of the specified range.
    - `SeriesLabels As Variant` (optional): True if the first row or column contains the name of the data series. False if the first row or column contains the first data point of the series. If this argument is omitted, Microsoft Excel attempts to determine the location of the series name from the contents of the first row or column.
    - `CategoryLabels As Variant` (optional): True if the first row or column contains the name of the category labels. False if the first row or column contains the first data point of the series. If this argument is omitted, Excel attempts to determine the location of the category label from the contents of the first row or column.
    - `Replace As Variant` (optional): If _CategoryLabels_ is True and _Replace_ is True, the specified categories replace the categories that currently exist for the series. If _Replace_ is False, the existing categories will not be replaced. The default value is False.
- `Extend(Source As Variant, [Rowcol As Variant], [CategoryLabels As Variant]) As Variant`  
  Adds new data points to an existing series collection.
    - `Source As Variant` (required): The new data to be added to the SeriesCollection object as a Range object.
    - `Rowcol As Variant` (optional): Specifies whether the new values are in the rows or columns of the given range source. Can be one of the following XlRowCol constants: xlRows or xlColumns. If this argument is omitted, Microsoft Excel attempts to determine where the values are by the size and orientation of the selected range or by the dimensions of the array.
    - `CategoryLabels As Variant` (optional): True to have the first row or column contain the name of the category labels. False to have the first row or column contain the first data point of the series. If this argument is omitted, Excel attempts to determine the location of the category label from the contents of the first row or column.
- `Item(Index As Variant) As Series`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
- `Paste([Rowcol As XlRowCol], [SeriesLabels As Variant], [CategoryLabels As Variant], [Replace As Variant], [NewSeries As Variant]) As Variant`  
  Pastes data from the Clipboard into the specified series collection.
    - `Rowcol As XlRowCol` (optional): Specifies whether the values corresponding to a particular data series are in rows or columns.
    - `SeriesLabels As Variant` (optional): True to use the contents of the cell in the first column of each row (or the first row of each column) as the name of the data series in that row (or column). False to use the contents of the cell in the first column of each row (or the first row of each column) as the first data point in the data series. The default value is False.
    - `CategoryLabels As Variant` (optional): True to use the contents of the first row (or column) of the selection as the categories for the chart. False to use the contents of the first row (or column) as the first data series in the chart. The default value is False.
    - `Replace As Variant` (optional): True to apply categories while replacing existing categories with information from the copied range. False to insert new categories without replacing any old ones. The default value is True.
    - `NewSeries As Variant` (optional): True to paste the data as a new series. False to paste the data as new points in an existing series. The default value is True.
- `NewSeries() As Series`  
  Creates a new series. Returns a Series object that represents the new series.
- `_Default(Index As Variant) As Series`
