# ThreeDFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209D0-0000-0000-C000-000000000046}  

Represents a shape's three-dimensional formatting.

**Remarks:** Use the ThreeD property to return a ThreeDFormat object. The following example adds an oval to the active document and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple. You cannot apply three-dimensional formatting to some kinds of shapes, such as beveled shapes or multiple-disjoint paths. Most of the properties and methods of the ThreeDFormat object for such a shape will fail.

## Properties (30)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ThreeDFormat object.
- `Depth As Single  (read/write)`  
  Returns or sets the depth of the shape's extrusion. Read/write Single.
- `ExtrusionColor As ColorFormat  (read-only)`  
  Returns a ColorFormat object that represents the color of the shape's extrusion. Read-only.
- `ExtrusionColorType As MsoExtrusionColorType  (read/write)`  
  Returns or sets a value that indicates whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. Read/write MsoExtrusionColorType.
- `Perspective As MsoTriState  (read/write)`  
  MsoTrue if the extrusion appears in perspective - that is, if the walls of the extrusion narrow toward a vanishing point. MsoFalse if the extrusion is a parallel, or orthographic, projection - that is, if the walls don't narrow toward a vanishing point. Read/write MsoTriState.
- `PresetExtrusionDirection As MsoPresetExtrusionDirection  (read-only)`  
  Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion). Read/write MsoPresetExtrusionDirection.
- `PresetLightingDirection As MsoPresetLightingDirection  (read/write)`  
  Returns or sets the position of the light source relative to the extrusion. Read/write MsoPresetLightingDirection.
- `PresetLightingSoftness As MsoPresetLightingSoftness  (read/write)`  
  Returns or sets the intensity of the extrusion lighting. Read/write MsoPresetLightingSoftness.
- `PresetMaterial As MsoPresetMaterial  (read/write)`  
  Returns or sets the extrusion surface material. Read/write MsoPresetMaterial.
- `PresetThreeDFormat As MsoPresetThreeDFormat  (read-only)`  
  Returns the preset extrusion format. Read-only MsoPresetThreeDFormat.
- `RotationX As Single  (read/write)`  
  Returns or sets the rotation of the extruded shape around the x-axis in degrees. Read/write Single.
- `RotationY As Single  (read/write)`  
  Returns or sets the rotation of the extruded shape around the y-axis, in degrees. Read/write Single.
- `Visible As MsoTriState  (read/write)`  
  True if the specified object, or the formatting applied to it, is visible. Read/write MsoTriState.
- `PresetLighting As MsoLightRigType  (read/write)`  
  Returns or sets an MsoBevelType constant that represents the lighting preset. Read/write.
- `Z As Single  (read/write)`  
  Returns or sets a Single that represents the z-axis for the shape. Read/write.
- `BevelTopType As MsoBevelType  (read/write)`  
  Returns or sets an MsoPresetCamera constant that represents the bevel type for the top bevel. Read/write.
- `BevelTopInset As Single  (read/write)`  
  Returns or sets a Single that represents the inset size for the top bevel. Read/write.
- `BevelTopDepth As Single  (read/write)`  
  Returns or sets a Single that represents the depth of the top bevel. Read/write.
- `BevelBottomType As MsoBevelType  (read/write)`  
  Returns or sets an MsoPresetCamera constant that represents the bevel type for the bottom bevel. Read/write.
- `BevelBottomInset As Single  (read/write)`  
  Returns or sets a Single that represents the inset size for the bottom bevel. Read/write.
- `BevelBottomDepth As Single  (read/write)`  
  Returns or sets a Single that represents the depth of the bottom bevel. Read/write.
- `PresetCamera As MsoPresetCamera  (read-only)`  
  Returns an MsoPresetCamera constant that represents the camera presets. Read-only.
- `RotationZ As Single  (read/write)`  
  Returns or sets a Single that represents z-axis rotation of the camera. Read/write.
- `ContourWidth As Single  (read/write)`  
  Returns or sets a Single that represents the width of the contour of a shape. Read/write.
- `ContourColor As ColorFormat  (read-only)`  
  Returns or sets a ColorFormat object that represents color of the contour of a shape. Read/write.
- `FieldOfView As Single  (read/write)`  
  Returns or sets a Single that represents the amount of perspective for a shape. Read/write.
- `ProjectText As MsoTriState  (read/write)`  
  Returns or sets an MsoTriState constant that represents whether text on a shape rotates with shape. msoTrue rotates the text. Read/write.
- `LightAngle As Single  (read/write)`  
  Returns or sets a Single that represents angle of the lighting. Read/write.

## Methods (9)

- `IncrementRotationX(Increment As Single)`  
  Changes the rotation of the specified shape around the x-axis by the specified number of degrees.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the shape around the x-axis is to be changed. Can be a value from -90 through 90. A positive value tilts the shape up; a negative value tilts it down.
- `IncrementRotationY(Increment As Single)`  
  Changes the rotation of the specified shape around the y-axis by the specified number of degrees.
    - `Increment As Single` (required): Specifies how much (in degrees) the rotation of the shape around the y-axis is to be changed. Can be a value from -90 through 90. A positive value tilts the shape to the left; a negative value tilts it to the right.
- `ResetRotation()`  
  Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward.
- `SetExtrusionDirection(PresetExtrusionDirection As MsoPresetExtrusionDirection)`  
  Sets the direction that the extrusion's sweep path takes away from the extruded shape.
    - `PresetExtrusionDirection As MsoPresetExtrusionDirection` (required): Sets the direction of the extrusion.
- `SetThreeDFormat(PresetThreeDFormat As MsoPresetThreeDFormat)`  
  Sets the preset extrusion format.
    - `PresetThreeDFormat As MsoPresetThreeDFormat` (required): Specifies a preset extrusion format that corresponds to one of the options (numbered from left to right, top to bottom) displayed when you click the 3D button on the Drawing toolbar.
- `SetPresetCamera(PresetCamera As MsoPresetCamera)`  
  Sets the camera presets for a shape.
    - `PresetCamera As MsoPresetCamera` (required): Specifies the camera preset type.
- `IncrementRotationZ(Increment As Single)`  
  Rotates a shape on the z-axis using the specified incrementation.
    - `Increment As Single` (required): Specifies the incrementation value.
- `IncrementRotationHorizontal(Increment As Single)`  
  Horizontally rotates a shape on the x-axis using the specified incrementation value.
    - `Increment As Single` (required): Specifies the incrementation value.
- `IncrementRotationVertical(Increment As Single)`  
  Vertically rotates a shape on the y-axis using the specified incrementation value.
    - `Increment As Single` (required): Specifies the incrementation value.
