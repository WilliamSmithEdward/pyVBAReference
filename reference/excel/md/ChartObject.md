# ChartObject

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208CF-0000-0000-C000-000000000046}  

Represents an embedded chart on a worksheet.

**Remarks:** The ChartObject object acts as a container for a Chart object. Properties and methods for the ChartObject object control the appearance and size of the embedded chart on the worksheet. The ChartObject object is a member of the ChartObjects collection. The ChartObjects collection contains all the embedded charts on a single sheet. Use ChartObjects (_index_), where _index_ is the embedded chart index number or name, to return a single ChartObject object.

**Example:**

```vba
Worksheets("Sheet1").ChartObjects(1).Chart. _
 ChartArea.Format.Fill.Pattern = msoPatternLightDownwardDiagonal
```

## Properties (21)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `BottomRightCell As Range  (read-only)`  
  Returns a Range object that represents the cell that lies under the lower-right corner of the object. Read-only.
- `Height As Double  (read/write)`  
  Returns or sets a Double value that represents the height, in points, of the object.
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `Left As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).
- `Locked As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if the object is locked.
- `Name As String  (read/write)`  
  Returns a String value representing the name of the object.
- `Placement As Variant  (read/write)`  
  Returns or sets a Variant value, containing an XlPlacement constant, that represents the way the object is attached to the cells below it.
- `PrintObject As Boolean  (read/write)`  
  True if the object will be printed when the document is printed. Read/write Boolean.
- `Top As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
- `TopLeftCell As Range  (read-only)`  
  Returns a Range object that represents the cell that lies under the upper-left corner of the specified object. Read-only.
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.
- `Width As Double  (read/write)`  
  Returns or sets a Double value that represents the width, in points, of the object.
- `ZOrder As Long  (read-only)`  
  Returns the z-order position of the object. Read-only Long.
- `ShapeRange As ShapeRange  (read-only)`  
  Returns a ShapeRange object that represents the specified object or objects. Read-only.
- `Chart As Chart  (read-only)`  
  Returns a Chart object that represents the chart contained in the object. Read-only.
- `ProtectChartObject As Boolean  (read/write)`  
  True if the embedded chart frame cannot be moved, resized, or deleted through the user interface. Read/write Boolean.
- `RoundedCorners As Boolean  (read/write)`  
  True if the embedded chart has rounded corners. Read/write Boolean.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if the font is a shadow font or if the object has a shadow.

## Methods (9)

- `BringToFront() As Variant`  
  Brings the object to the front of the z-order.
- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat]) As Variant`  
  Copies the selected object to the Clipboard as a picture.
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
- `SendToBack() As Variant`  
  Sends the object to the back of the z-order.
- `Activate() As Variant`  
  Makes the current chart the active chart.
- `Copy() As Variant`  
  Copies the object to the Clipboard.
