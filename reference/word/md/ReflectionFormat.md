# ReflectionFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {F01943FF-1985-445E-8602-8FB8F39CCA75}  

Represents the reflection formatting for a shape or range of shapes.

## Properties (8)

- `Type As MsoReflectionType  (read/write)`  
  Returns or sets an MsoLightRigType constant that represents the type and direction of the lighting for a shape reflection. Read/write.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ReflectionFormat object.
- `Transparency As Single  (read/write)`  
  Returns or sets the degree of transparency for the reflection effect as a value between 0.0 (opaque) and 1.0 (clear). Read/write Single.
- `Size As Single  (read/write)`  
  Returns or sets the size of the reflection as a percentage of the reflected shape from 0 to 100. Read/write.
- `Offset As Single  (read/write)`  
  Returns or sets the amount of separation, in points, of the reflected image from the shape. Read/write.
- `Blur As Single  (read/write)`  
  Returns or sets a Single that specifies the degree of blur effect applied to the specified object. Read/write.
