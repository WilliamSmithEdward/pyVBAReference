# Model3DFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C03D8-0000-0000-C000-000000000046}  

Represents the properties of a 3D model shape.

## Properties (14)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `AutoFit As MsoTriState  (read/write)`  
  Returns whether AutoFit is enabled for the model. Read/write.
- `RotationX As Single  (read/write)`  
  Returns the x-angle of a 3D model object's rotation. Read/write.
- `RotationY As Single  (read/write)`  
  Returns the y-angle of a 3D model object's rotation. Read/write.
- `RotationZ As Single  (read/write)`  
  Returns the z-angle of a 3D model object's rotation. Read/write.
- `FieldOfView As Single  (read/write)`  
  Returns the field-of-view angle of a 3D model object's camera, expressed in degrees. Read/write.
- `CameraPositionX As Single  (read/write)`  
  Returns the x-coordinate of a 3D model object's camera position. Read/write.
- `CameraPositionY As Single  (read/write)`  
  Returns the y-coordinate of a 3D model object's camera position. Read/write.
- `CameraPositionZ As Single  (read/write)`  
  Returns the z-coordinate of a 3D model object's camera position. Read/write.
- `LookAtPointX As Single  (read/write)`  
  Returns the x-coordinate of a 3D model object's camera look-at position. Read/write.
- `LookAtPointY As Single  (read/write)`  
  Returns the y-coordinate of a 3D model object's camera look-at position. Read/write.
- `LookAtPointZ As Single  (read/write)`  
  Returns the z-coordinate of a 3D model object's camera look-at position. Read/write.

## Methods (4)

- `ResetModel([ResetSize As Boolean])`  
  Changes the rotation of the specified shape around the x-axis by the specified number of degrees.
    - `ResetSize As Boolean` (optional): True to reset the 3D model frame to the same size as when a model is first inserted; False to leave the 3D model frame size alone.
- `IncrementRotationX(Increment As Single)`  
  Changes the rotation of the specified shape around the x-axis by the specified number of degrees.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the model around the x-axis is to be changed. Any value can be provided, although any value will be effectively normalized into the range 0..360 degrees.
- `IncrementRotationY(Increment As Single)`  
  Changes the rotation of the specified shape around the y-axis by the specified number of degrees.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the model around the y-axis is to be changed. Any value can be provided, although any value will be effectively normalized into the range 0..360 degrees.
- `IncrementRotationZ(Increment As Single)`  
  Changes the rotation of the specified shape around the z-axis by the specified number of degrees.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the model around the z-axis is to be changed. Any value can be provided, although any value will be effectively normalized into the range 0..360 degrees.
