# ThreeDFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C0321-0000-0000-C000-000000000046}  

Represents a shape's three-dimensional formatting.

**Remarks:** You cannot apply three-dimensional formatting to some kinds of shapes, such as beveled shapes or multiple-disjoint paths. Most of the properties and methods of the ThreeDFormat object for such a shape will fail.

**Example:**

```vba
Set myDocument = Worksheets(1)
Set myShape = myDocument.Shapes.AddShape(msoShapeOval, _
 90, 90, 90, 40)
With myShape.ThreeD
 .Visible = True
 .Depth = 50
 .ExtrusionColor.RGB = RGB(255, 100, 255)
 ' RGB value for purple
End With
```

## Properties (30)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Depth As Single  (read/write)`  
  Returns or sets a Single value that represents the depth of the shape's extrusion.
- `ExtrusionColor As ColorFormat  (read-only)`  
  Returns a ColorFormat object that represents the color of the shape's extrusion. Read-only.
- `ExtrusionColorType As MsoExtrusionColorType  (read/write)`  
  Returns or sets a value that indicates whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. Read/write MsoExtrusionColorType.
- `Perspective As MsoTriState  (read/write)`  
  Returns or sets an MsoTriState value that determines whether the extrusion appears in perspective.
- `PresetExtrusionDirection As MsoPresetExtrusionDirection  (read-only)`  
  Returns the direction that the extrusion's sweep path takes away from the extruded shape (the front face of the extrusion). Read-only MsoPresetExtrusionDirection.
- `PresetLightingDirection As MsoPresetLightingDirection  (read/write)`  
  Returns or sets the position of the light source relative to the extrusion. Read/write MsoPresetLightingDirection.
- `PresetLightingSoftness As MsoPresetLightingSoftness  (read/write)`  
  Returns or sets the intensity of the extrusion lighting. Read/write MsoPresetLightingSoftness.
- `PresetMaterial As MsoPresetMaterial  (read/write)`  
  Returns or sets the extrusion surface material. Read/write MsoPresetMaterial.
- `PresetThreeDFormat As MsoPresetThreeDFormat  (read-only)`  
  Returns the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion. Read-only MsoPresetThreeDFormat.
- `RotationX As Single  (read/write)`  
  Returns or sets the rotation of the extruded shape around the x-axis in degrees. Can be a value from -90 through 90. A positive value indicates upward rotation; a negative value indicates downward rotation. Read/write Single.
- `RotationY As Single  (read/write)`  
  Returns or sets the rotation of the extruded shape around the y-axis in degrees. Can be a value from -90 through 90. A positive value indicates rotation to the left; a negative value indicates rotation to the right. Read/write Single.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets an MsoTriState value that determines whether the object is visible. Read/write.
- `PresetLighting As MsoLightRigType  (read/write)`  
  Returns or sets the extrusion preset lighting for a ThreeDFormat object. Read-only MsoLightRigType.
- `Z As Single  (read/write)`  
  Returns the Z order of the specified ThreeDFormat object. Read/write Single.
- `BevelTopType As MsoBevelType  (read/write)`  
  Returns or sets the top bevel type for a ThreeDFormat object. Read/write MsoBevelType.
- `BevelTopInset As Single  (read/write)`  
  Returns or sets a value indicating whether the top insert bevel should be raised for a ThreeDFormat object. Read/write Single.
- `BevelTopDepth As Single  (read/write)`  
  Returns or sets the top depth when using the bevel effect on a ThreeDFormat object. Read/write Single.
- `BevelBottomType As MsoBevelType  (read/write)`  
  Returns or sets the bottom bevel type for a ThreeDFormat object. Read/write MsoBevelType.
- `BevelBottomInset As Single  (read/write)`  
  Returns or sets a value indicating whether the bottom insert bevel should be raised for a ThreeDFormat object. Read/write Single.
- `BevelBottomDepth As Single  (read/write)`  
  Returns or sets the bottom depth when using the bevel effect on a ThreeDFormat object. Read/write Single.
- `PresetCamera As MsoPresetCamera  (read-only)`  
  Returns or sets the extrusion preset camera for a ThreeDFormat object. Read-only MsoPresetCamera.
- `RotationZ As Single  (read/write)`  
  Returns or sets the rotation of the extruded shape around the z-axis in degrees. Can be a value from -90 through 90. A positive value indicates upward rotation; a negative value indicates downward rotation. Read/write Single.
- `ContourWidth As Single  (read/write)`  
  Returns or sets the contour width for a ThreeDFormat object. Read/write Single.
- `ContourColor As ColorFormat  (read-only)`  
  Returns the contour color for a ThreeDFormat object. Read-only ColorFormat.
- `FieldOfView As Single  (read/write)`  
  Returns or sets the angle at which a ThreeDFormat object can be viewed. Read/write Single.
- `ProjectText As MsoTriState  (read/write)`  
  Returns or sets the project text state for the specified ThreeDFormat object. Read/write MsoTriState.
- `LightAngle As Single  (read/write)`  
  Returns or sets the angle of the extrusion lights set on a ThreeDFormat object. Read/write Single.

## Methods (9)

- `IncrementRotationX(Increment As Single)`  
  Changes the rotation of the specified shape around the x-axis by the specified number of degrees. Use the RotationX property to set the absolute rotation of the shape around the x-axis.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the shape around the x-axis is to be changed. Can be a value from -90 through 90. A positive value tilts the shape up; a negative value tilts it down.
- `IncrementRotationY(Increment As Single)`  
  Changes the rotation of the specified shape around the y-axis by the specified number of degrees. Use the RotationY property to set the absolute rotation of the shape around the y-axis.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the shape around the y-axis is to be changed. Can be a value from -90 through 90. A positive value tilts the shape to the left; a negative value tilts it to the right.
- `ResetRotation()`  
  Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.
- `SetThreeDFormat(PresetThreeDFormat As MsoPresetThreeDFormat)`  
  Sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.
    - `PresetThreeDFormat As MsoPresetThreeDFormat` (required): Specifies a preset extrusion format that corresponds to one of the options (numbered from left to right, from top to bottom) displayed when you choose the 3D button on the Drawing toolbar.
- `SetExtrusionDirection(PresetExtrusionDirection As MsoPresetExtrusionDirection)`  
  Sets the direction that the extrusion's sweep path takes away from the extruded shape.
    - `PresetExtrusionDirection As MsoPresetExtrusionDirection` (required): Specifies the extrusion direction.
- `SetPresetCamera(PresetCamera As MsoPresetCamera)`  
  Sets the camera for the specified ThreeDFormat object.
    - `PresetCamera As MsoPresetCamera` (required): Specifies the preset camera.
- `IncrementRotationZ(Increment As Single)`  
  Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the RotationZ property to set the absolute rotation of the shape around the z-axis.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the shape around the z-axis is to be changed. Can be a value from -90 through 90. A positive value tilts the shape to the left; a negative value tilts it to the right.
- `IncrementRotationHorizontal(Increment As Single)`  
  Changes the rotation of the specified shape horizontally by the specified number of degrees.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the shape is to be changed horizontally. Can be a value from -90 through 90. A positive value moves the shape left; a negative value moves it to the right.
- `IncrementRotationVertical(Increment As Single)`  
  Changes the rotation of the specified shape vertically by the specified number of degrees.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the shape is to be changed vertically. Can be a value from -90 through 90. A positive value tilts the shape up; a negative value tilts it down.
