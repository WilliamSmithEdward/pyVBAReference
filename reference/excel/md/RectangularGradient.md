# RectangularGradient

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244B0-0000-0000-C000-000000000046}  

The RectangularGradient object transitions through a series of colors in a linear manner along a specific angle.

**Remarks:** Attempting to access a Gradient property of an Interior object that does not have an existing gradient fill results in a run-time error. Be aware of the Pattern property of the Interior object before accessing the Gradient property. If the Pattern property is changed from a gradient type to a non-gradient type, the Gradient property will populate with default values.

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ColorStops As ColorStops  (read-only)`  
  Returns the ColorStops collection for the RectangularGradient object. Read-only.
- `RectangleTop As Double  (read/write)`  
  Represents the point or vector that the gradient fill converges to. Read/write.
- `RectangleBottom As Double  (read/write)`  
  Represents the point or vector that the gradient fill converges to. Read/write.
- `RectangleLeft As Double  (read/write)`  
  Represents the point or vector that the gradient fill converges to. Read/write.
- `RectangleRight As Double  (read/write)`  
  Represents the point or vector that the gradient fill converges to. Read/write.
