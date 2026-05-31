# ChartObjects

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208D0-0000-0000-C000-000000000046}  

A collection of all the ChartObject objects on the specified chart sheet, dialog sheet, or worksheet.

**Remarks:** Each ChartObject object represents an embedded chart. The ChartObject object acts as a container for a Chart object. Properties and methods for the ChartObject object control the appearance and size of the embedded chart on the sheet.

**Example:**

```vba
Worksheets("sheet1").ChartObjects.Delete
```

## Properties (14)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Height As Double  (read/write)`  
  Returns or sets a Double value that represents the height, in points, of the object.
- `Left As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).
- `Locked As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if the objects are locked.
- `Placement As Variant  (read/write)`  
  Returns or sets a Variant value, containing an XlPlacement constant, that represents the way the objects are attached to the cells below them.
- `PrintObject As Boolean  (read/write)`  
  True if the objects will be printed when the document is printed. Read/write Boolean.
- `Top As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.
- `Width As Double  (read/write)`  
  Returns or sets a Double value that represents the width, in points, of the object.
- `ShapeRange As ShapeRange  (read-only)`  
  Returns a ShapeRange object that represents the specified object or objects. Read-only.
- `ProtectChartObject As Boolean  (read/write)`  
  True if the embedded chart frame cannot be moved, resized, or deleted through the user interface. Read/write Boolean.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (10)

- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat]) As Variant`  
  Copies the selected object to the Clipboard as a picture. Variant.
    - `Appearance As XlPictureAppearance` (optional): Specifies how the picture should be copied. The default value is xlScreen.
    - `Format As XlCopyPictureFormat` (optional): The format of the picture. The default value is xlPicture.
- `Cut() As Variant`  
  Cuts the object to the Clipboard.
- `Delete() As Variant`  
  Deletes the object.
- `Duplicate() As Object`  
  Duplicates the object and returns a reference to the new copy.
- `Select([Replace As Variant]) As Variant`  
  Selects the object.
    - `Replace As Variant` (optional): True to replace the current selection with the specified object. False to extend the current selection to include any previously selected objects and the specified object.
- `Copy() As Variant`  
  Copies the object to the Clipboard.
- `Add(Left As Double, Top As Double, Width As Double, Height As Double) As ChartObject`  
  Creates a new embedded chart.
    - `Left As Double` (required): The initial coordinates of the new object (in points), relative to the upper-left corner of cell A1 on a worksheet or to the upper-left corner of a chart.
    - `Width As Double` (required): The initial size of the new object, in points.
- `Item(Index As Variant) As Object`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
- `_Default(Index As Variant) As Object`
