# ChartFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A5C-F07E-4CA4-AF6F-BEF486AA4E6F}  

Provides access to the OfficeArt formatting for chart elements.

**Remarks:** Using a property or method that does not apply to the type of object the ChartFormat object is attached to will result in a run-time error.

## Properties (13)

- `Fill As FillFormat  (read-only)`  
  Returns a FillFormat object for the parent chart element that contains fill formatting properties for the chart element. Read-only.
- `Glow As GlowFormat  (read-only)`  
  Returns the glow formatting properties for the chart element. Read-only GlowFormat.
- `Line As LineFormat  (read-only)`  
  Returns the line formatting properties for the specified chart element. Read-only LineFormat.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `PictureFormat As PictureFormat  (read-only)`  
  Returns a PictureFormat object for a specified chart that contains pictures. Read-only.
- `Shadow As ShadowFormat  (read-only)`  
  Returns shadow formatting properties for the chart element. Read-only ShadowFormat.
- `SoftEdge As SoftEdgeFormat  (read-only)`  
  Returns the soft edge formatting for a shape. Read-only SoftEdgeFormat.
- `TextFrame2 As TextFrame2  (read-only)`  
  Returns the text formatting for the specified chart element. Read-only TextFrame2.
- `ThreeD As ThreeDFormat  (read-only)`  
  Returns the 3D-effect formatting properties for the specified chart. Read-only ThreeDFormat.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Adjustments As Adjustments  (read-only)`  
  Returns an Adjustments object that contains a collection of adjustment values for the specified Chart Format object.
- `AutoShapeType As MsoAutoShapeType  (read/write)`  
  Returns or sets the auto shape type of the specified object. Read/write MSOAutoShapeType.
