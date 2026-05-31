# OLEObject

**Type:** Class  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020818-0000-0000-C000-000000000046}  

Represents an ActiveX control or a linked or embedded OLE object on a worksheet.

**Remarks:** The OLEObject object is a member of the OLEObjects collection. The OLEObjects collection contains all the OLE objects on a single worksheet.

**Example:**

```vba
Worksheets("sheet1").OLEObjects(1).Delete
```

## Properties (29)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `BottomRightCell As Range  (read-only)`  
  Returns a Range object that represents the cell that lies under the lower-right corner of the object. Read-only.
- `Enabled As Boolean  (read/write)`  
  True if the object is enabled. Read/write Boolean.
- `Height As Double  (read/write)`  
  Returns or sets a Double value that represents the height, in points, of the object.
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `Left As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).
- `Locked As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if the object is locked.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `Placement As Variant  (read/write)`  
  Returns or sets a Variant value containing an XlPlacement constant that represents the way the object is attached to the cells below it.
- `PrintObject As Boolean  (read/write)`  
  True if the object is printed when the document is printed. Read/write Boolean.
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
- `Border As Border  (read-only)`  
  Returns a Border object that represents the border of the object.
- `Interior As Interior  (read-only)`  
  Returns an Interior object that represents the interior of the specified object.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if the object has a shadow.
- `AutoLoad As Boolean  (read/write)`  
  True if the OLE object is automatically loaded when the workbook that contains it is opened. Read/write Boolean.
- `AutoUpdate As Boolean  (read/write)`  
  True if the OLE object is updated automatically when the source changes. Valid only if the object is linked; its OLEType property must be xlOLELink (XlOLEType enumeration). Read-only Boolean.
- `Object As Object  (read-only)`  
  Returns the OLE Automation object associated with this OLE object. Read-only Object.
- `OLEType As Variant  (read-only)`  
  Returns the OLE object type. Can be one of the following XlOLEType constants: xlOLELink or xlOLEEmbed. Returns xlOLELink if the object is linked (it exists outside of the file), or returns xlOLEEmbed if the object is embedded (it's entirely contained within the file). Read-only Long.
- `SourceName As String  (read/write)`  
  Returns or sets a String value that represents the specified object's link source name.
- `LinkedCell As String  (read/write)`  
  Returns or sets the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell. Read/write String.
- `ListFillRange As String  (read/write)`  
  Returns or sets the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box. Read/write String.
- `progID As String  (read-only)`  
  Returns the programmatic identifiers for the object. Read-only String.

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
  Cuts the object to the Clipboard or pastes it into a specified destination.
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
  Activates the object.
- `Update() As Variant`  
  Updates the link.
- `Verb([Verb As XlOLEVerb]) As Variant`  
  Sends a verb to the server of the specified OLE object.
    - `Verb As XlOLEVerb` (optional): The verb that the server of the OLE object should act on. If this argument is omitted, the default verb is sent. The available verbs are determined by the object's source application. Typical verbs for an OLE object are Open and Primary (represented by the XlOLEVerb constants xlOpen and xlPrimary).

## Events (2)

- `GotFocus()`  
  Occurs when an ActiveX control gets input focus.
- `LostFocus()`  
  Occurs when an ActiveX control loses input focus.
