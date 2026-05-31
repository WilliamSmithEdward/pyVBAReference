# CalloutFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C0311-0000-0000-C000-000000000046}  

Contains properties and methods that apply to line callouts.

**Remarks:** Use the Callout property of the Shape object to return a CalloutFormat object.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes(3).Callout
 .Accent = True
 .Angle = msoCalloutAngle30
 .Border = False
 .PresetDrop msoCalloutDropTop
 .Type = msoCalloutThree
End With
```

## Properties (13)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Accent As MsoTriState  (read/write)`  
  Allows the user to place a vertical accent bar to separate the callout text from the callout line. Read/write MsoTriState.
- `Angle As MsoCalloutAngleType  (read/write)`  
  Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write MsoCalloutAngleType.
- `AutoAttach As MsoTriState  (read/write)`  
  True if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box. Read/write MsoTriState.
- `AutoLength As MsoTriState  (read-only)`  
  Applies only to callouts whose lines consist of more than one segment (MsoCalloutType types msoCalloutThree and msoCalloutFour). Read/write MsoTriState.
- `Border As MsoTriState  (read/write)`  
  Returns or sets an MsoTriState value that represents the visibility options for the border of the object.
- `Drop As Single  (read-only)`  
  For callouts with an explicitly set drop value, this property returns the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. Read-only Single.
- `DropType As MsoCalloutDropType  (read-only)`  
  Returns a value that indicates where the callout line attaches to the callout text box. Read-only MsoCalloutDropType.
- `Gap As Single  (read/write)`  
  Returns or sets the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write Single.
- `Length As Single  (read-only)`  
  Returns a Single value that represents the length (in points) of the first segment of the callout line (the segment attached to the text callout box).
- `Type As MsoCalloutType  (read/write)`  
  Returns or sets an MsoCalloutType value that represents the callout format type.

## Methods (4)

- `AutomaticLength()`  
  Specifies that the first segment of the callout line (the segment attached to the text callout box) be scaled automatically when the callout is moved.
- `CustomDrop(Drop As Single)`  
  Sets the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box.
    - `Drop As Single` (required): The drop distance, in points.
- `CustomLength(Length As Single)`  
  Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved.
    - `Length As Single` (required): The length of the first segment of the callout, in points.
- `PresetDrop(DropType As MsoCalloutDropType)`  
  Specifies whether the callout line attaches to the top, bottom, or center of the callout text box, or whether it attaches at a point that's a specified distance from the top or bottom of the text box.
    - `DropType As MsoCalloutDropType` (required): The starting position of the callout line relative to the text bounding box.
