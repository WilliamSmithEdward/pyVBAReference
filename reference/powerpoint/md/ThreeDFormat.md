# ThreeDFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493483-5A91-11CF-8700-00AA0060263B}  

Represents a shape's three-dimensional formatting.

**Remarks:** You cannot apply three-dimensional formatting to some kinds of shapes, such as beveled shapes or multiple-disjoint paths. Most of the properties and methods of the ThreeDFormat object for such a shape will fail.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

Set myShape = myDocument.Shapes _
    .AddShape(msoShapeOval, 90, 90, 90, 40)

With myShape.ThreeD
    .Visible = True
    .Depth = 50
    'RGB value for purple
    .ExtrusionColor.RGB = RGB(255, 100, 255)
End With
```

## Properties (30)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Depth As Single  (read/write)`  
  Returns or sets the depth of the shape's extrusion. Read/write.
- `ExtrusionColor As ColorFormat  (read-only)`  
  Returns a ColorFormat object that represents the color of the shape's extrusion. Read-only.
- `ExtrusionColorType As MsoExtrusionColorType  (read/write)`  
  Returns or sets a value that indicates whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. Read/write.
- `Perspective As MsoTriState  (read/write)`  
  Determines whether the extrusion appears in perspective. Read/write.
- `PresetExtrusionDirection As MsoPresetExtrusionDirection  (read-only)`  
  Returns the direction that the extrusion's sweep path takes away from the extruded shape (the front face of the extrusion). Read-only.
- `PresetLightingDirection As MsoPresetLightingDirection  (read/write)`  
  Returns or sets the position of the light source relative to the extrusion. Read/write.
- `PresetLightingSoftness As MsoPresetLightingSoftness  (read/write)`  
  Returns or sets the intensity of the extrusion lighting. Read/write.
- `PresetMaterial As MsoPresetMaterial  (read/write)`  
  Returns or sets the extrusion surface material. Read/write.
- `PresetThreeDFormat As MsoPresetThreeDFormat  (read-only)`  
  Returns the preset extrusion format. Read-only.
- `RotationX As Single  (read/write)`  
  Returns or sets the rotation of the extruded shape around the x-axis, in degrees. Read/write.
- `RotationY As Single  (read/write)`  
  Returns or sets the rotation of the extruded shape around the y-axis, in degrees. Read/write.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.
- `PresetLighting As MsoLightRigType  (read/write)`  
  Returns or sets the effects lighting used by the specified ThreeDFormat object. Read/write.
- `Z As Single  (read/write)`  
  Returns or sets the distance the specified ThreeDFormat object is moved away from the ground plane, in points. Read/write.
- `BevelTopType As MsoBevelType  (read/write)`  
  Returns or set the top bevel type. Read/write.
- `BevelTopInset As Single  (read/write)`  
  Returns or sets the top bevel inset for the specified ThreeDFormat object, in points. Read/write.
- `BevelTopDepth As Single  (read/write)`  
  Returns or sets the top bevel height for the specified ThreeDFormat object, in points. Read/write.
- `BevelBottomType As MsoBevelType  (read/write)`  
  Returns or set the bottom bevel type. Read/write.
- `BevelBottomInset As Single  (read/write)`  
  Returns or sets the bottom bevel inset for the specified ThreeDFormat object, in points. Read/write.
- `BevelBottomDepth As Single  (read/write)`  
  Returns or sets the bottom bevel height for the specified ThreeDFormat object, in points. Read/write.
- `PresetCamera As MsoPresetCamera  (read-only)`  
  Returns the effects camera type used by the specified ThreeDFormat object. Read-only.
- `RotationZ As Single  (read/write)`  
  Returns or sets the rotation of the effects camera for the specified ThreeDFormat object, in degrees. Read/write.
- `ContourWidth As Single  (read/write)`  
  Returns or sets the width of the contour around the specified ThreeDFormat object, in points. Read/write.
- `ContourColor As ColorFormat  (read-only)`  
  Returns a ColorFormat object that represents the color of the specified ThreeDFormat object's contour. Read-only.
- `FieldOfView As Single  (read/write)`  
  Returns or sets the camera field of view for the specified ThreeDFormat object, in degrees. Read/write.
- `ProjectText As MsoTriState  (read/write)`  
  Specifies whether text on the specified ThreeDFormat object rotates with the object. Read/write.
- `LightAngle As Single  (read/write)`  
  Sets or returns the angle at which light impacts the shape whose three-dimensional format is represented by the specified object. Read/write.

## Methods (9)

- `IncrementRotationX(Increment As Single)`  
  Changes the rotation of the specified shape around the x-axis by the specified number of degrees.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the shape around the x-axis is to be changed. Can be a value from -90 through 90. A positive value tilts the shape up; a negative value tilts it down.
- `IncrementRotationY(Increment As Single)`  
  Changes the rotation of the specified shape around the y-axis by the specified number of degrees.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the shape around the y-axis is to be changed. Can be a value from -90 through 90. A positive value tilts the shape to the left; a negative value tilts it to the right.
- `ResetRotation()`  
  Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.
- `SetThreeDFormat(PresetThreeDFormat As MsoPresetThreeDFormat)`  
  Sets the preset extrusion format.
    - `PresetThreeDFormat As MsoPresetThreeDFormat` (required): Specifies a preset extrusion format that corresponds to one of the options (numbered from left to right, from top to bottom) displayed when you click the 3D Rotation submenu on the Shape Effects menu.
- `SetExtrusionDirection(PresetExtrusionDirection As MsoPresetExtrusionDirection)`  
  Sets the direction that the extrusion's sweep path takes away from the extruded shape.
    - `PresetExtrusionDirection As MsoPresetExtrusionDirection` (required): Specifies the extrusion direction.
- `SetPresetCamera(PresetCamera As MsoPresetCamera)`  
  Specifies the effects camera type and rotation to use for the specified ThreeDFormat object.
    - `PresetCamera As MsoPresetCamera` (required): The ThreeDFormat object's effects camera type and rotation.
- `IncrementRotationZ(Increment As Single)`  
  Rotates the ThreeDFormat object around the ground plane by the specified number of degrees.
    - `Increment As Single` (required): The number of degrees to rotate the ThreeDFormat object, between -180 and 180 degrees.
- `IncrementRotationHorizontal(Increment As Single)`  
  Rotates the ThreeDFormat object along the horizontal axis by the specified number of degrees.
    - `Increment As Single` (required): The number of degrees to rotate the ThreeDFormat object, between -180 and 180 degrees.
- `IncrementRotationVertical(Increment As Single)`  
  Rotates the ThreeDFormat object along the vertical axis by the specified number of degrees.
    - `Increment As Single` (required): The number of degrees to rotate the ThreeDFormat object, between -180 and 180 degrees.
