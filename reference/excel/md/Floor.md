# Floor

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208C7-0000-0000-C000-000000000046}  

Represents the floor of a 3D chart.

**Example:**

```vba
Worksheets("sheet1").ChartObjects(1).Activate
ActiveChart.Floor.Interior.Color = RGB(0, 255, 255)
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `PictureType As Variant  (read/write)`  
  Returns or sets a Variant value that represents the way pictures are displayed on the walls and faces of a 3D chart.
- `Thickness As Long  (read/write)`  
  Returns or sets a Long, specifying the thickness of the floor. Read/write.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (3)

- `Select() As Variant`  
  Selects the object.
- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `Paste()`  
  Pastes a picture from the Clipboard on the floor of the specified chart.
