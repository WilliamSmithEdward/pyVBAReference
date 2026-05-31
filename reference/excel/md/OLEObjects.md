# OLEObjects

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208A3-0000-0000-C000-000000000046}  

A collection of all the OLEObject objects on the specified worksheet.

**Remarks:** Each OLEObject object represents an ActiveX control or a linked or embedded OLE object. An ActiveX control on a sheet has two names: the name of the shape that contains the control, which you can see in the Name box when you view the sheet, and the code name for the control, which you can see in the cell to the right of (Name) in the Properties window. When you first add a control to a sheet, the shape name and code name match. However, if you change either the shape name or code name, the other is not automatically changed to match. The latter however, seems to have changed with Excel versions. With version 16.0 both are kept consistent and is not possible to change one of the two alone.

**Example:**

```vba
Worksheets(1).OLEObjects.Visible = False
```

## Properties (20)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Enabled As Boolean  (read/write)`  
  True if the object is enabled. Read/write Boolean.
- `Height As Double  (read/write)`  
  Returns or sets a Double value that represents the height, in points, of the object.
- `Left As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).
- `Locked As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if the object is locked.
- `Placement As Variant  (read/write)`  
  Returns or sets a Variant value containing an XlPlacement constant that represents the way the object is attached to the cells below it.
- `PrintObject As Boolean  (read/write)`  
  True if the object will be printed when the document is printed. Read/write Boolean.
- `Top As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.
- `Width As Double  (read/write)`  
  Returns or sets a Double value that represents the width, in points, of the object.
- `ZOrder As Long  (read-only)`  
  Returns the z-order position of the object. Read-only Long.
- `ShapeRange As ShapeRange  (read-only)`  
  Returns a ShapeRange object that represents the specified object or objects. Read-only.
- `Border As Border  (read-only)`  
  Returns a Border object that represents the border of the object.
- `Interior As Interior  (read-only)`  
  Returns an Interior object that represents the interior of the specified object.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if the object has a shadow.
- `AutoLoad As Boolean  (read/write)`  
  True if the OLE object is automatically loaded when the workbook that contains it is opened. Read/write Boolean.
- `SourceName As String  (read/write)`  
  Returns or sets a String value that represents the specified object's link source name.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (11)

- `BringToFront() As Variant`  
  Brings the object to the front of the z-order.
- `Copy() As Variant`  
  Copies the object to the Clipboard.
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
- `SendToBack() As Variant`  
  Sends the object to the back of the z-order.
- `Add([ClassType As Variant], [Filename As Variant], [Link As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As OLEObject`  
  Adds a new OLE object to a sheet.
    - `ClassType As Variant` (optional): You must specify either _ClassType_ or _FileName_. A string that contains the programmatic identifier for the object to be created. If _ClassType_ is specified, _FileName_ and _Link_ are ignored.
    - `Filename As Variant` (optional): You must specify either _ClassType_ or _FileName_. A string that specifies the file to be used to create the OLE object.
    - `Link As Variant` (optional): True to have the new OLE object based on _FileName_ be linked to that file. If the object isn't linked, the object is created as a copy of the file. The default value is False.
    - `DisplayAsIcon As Variant` (optional): True to display the new OLE object either as an icon or as its regular picture. If this argument is True, _IconFileName_ and _IconIndex_ can be used to specify an icon.
    - `IconFileName As Variant` (optional): A string that specifies the file that contains the icon to be displayed. This argument is used only if _DisplayAsIcon_ is True. If this argument isn't specified or the file contains no icons, the default icon for the OLE class is used.
    - `IconIndex As Variant` (optional): The number of the icon in the icon file. This is used only if _DisplayAsIcon_ is True and _IconFileName_ refers to a valid file that contains icons. If an icon with the given index number doesn't exist in the file specified by _IconFileName_, the first icon in the file is used.
    - `IconLabel As Variant` (optional): A string that specifies a label to display beneath the icon. This is used only if _DisplayAsIcon_ is True. If this argument is omitted or is an empty string (""), no caption is displayed.
    - `Left As Variant` (optional): The initial coordinates of the new object, in points, relative to the upper-left corner of cell A1 on a worksheet, or to the upper-left corner of a chart.
    - `Top As Variant` (optional): The initial coordinates of the new object, in points, relative to the top of row 1 on a worksheet, or to the top of the chart area on a chart.
    - `Width As Variant` (optional): The initial width of the new object, in points.
    - `Height As Variant` (optional): The initial height of the new object, in points.
- `Item(Index As Variant) As Object`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
