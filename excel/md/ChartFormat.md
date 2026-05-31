# ChartFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244B2-0000-0000-C000-000000000046}  

Provides access to the Office Art formatting for chart elements.

**Remarks:** Using a property or method that does not apply to the type of object that the ChartFormat object is attached to will result in a run-time error.

## Properties (13)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Fill As FillFormat  (read-only)`  
  Returns a FillFormat object for the parent chart element that contains fill formatting properties for the chart element. Read-only.
- `Glow As GlowFormat  (read-only)`  
  Returns a GlowFormat object for a specified chart that contains glow formatting properties for the chart element. Read-only.
- `Line As LineFormat  (read-only)`  
  Returns a LineFormat object that contains line formatting properties for the specified chart element. Read-only.
- `PictureFormat As PictureFormat  (read-only)`  
  Returns a PictureFormat object for a specified chart that contains pictures. Read-only.
- `Shadow As ShadowFormat  (read-only)`  
  Returns a ShadowFormat object that contains shadow formatting properties for the chart element. Read-only.
- `SoftEdge As SoftEdgeFormat  (read-only)`  
  Returns a SoftEdgeFormat object for a specified chart that contains soft edge formatting properties for the chart. Read-only.
- `TextFrame2 As TextFrame2  (read-only)`  
  Returns a TextFrame2 object that contains text formatting for the specified chart element. Read-only.
- `ThreeD As ThreeDFormat  (read-only)`  
  Returns a ThreeDFormat object that contains 3D-effect formatting properties for the specified chart. Read-only.
- `Adjustments As Adjustments  (read-only)`  
  Returns an Adjustments object. Read-only.
- `AutoShapeType As MsoAutoShapeType  (read/write)`  
  Returns the type of the specified shape. Read-only MsoAutoShapeType.
