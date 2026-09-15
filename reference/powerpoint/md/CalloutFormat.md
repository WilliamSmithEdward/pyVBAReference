# CalloutFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493485-5A91-11CF-8700-00AA0060263B}  

Contains properties and methods that apply to line callouts.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

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
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Accent As MsoTriState  (read/write)`  
  Determines whether a vertical accent bar separates the callout text from the callout line. Read/write.
- `Angle As MsoCalloutAngleType  (read/write)`  
  Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write.
- `AutoAttach As MsoTriState  (read/write)`  
  Determines whether the place where the callout line attaches to the callout text box changes, depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box. Read/write.
- `AutoLength As MsoTriState  (read-only)`  
  Determines whether the first segment of the callout retains the fixed length specified by the Length property, or is scaled automatically, whenever the callout is moved. Read-only.
- `Border As MsoTriState  (read/write)`  
  Determines whether the text in the specified callout is surrounded by a border. Read/write.
- `Drop As Single  (read-only)`  
  For callouts with an explicitly set drop value, this property returns the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. Read-only.
- `DropType As MsoCalloutDropType  (read-only)`  
  Returns a value that indicates where the callout line attaches to the callout text box. Read-only.
- `Gap As Single  (read/write)`  
  Returns or sets the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write.
- `Length As Single  (read-only)`  
  When the AutoLength property of the specified callout is set to False, the Length property returns the length (in points) of the first segment of the callout line (the segment attached to the text callout box). Read-only.
- `Type As MsoCalloutType  (read/write)`  
  Represents the type of callout. Read/write.

## Methods (4)

- `AutomaticLength()`  
  Specifies that the first segment of the callout line (the segment attached to the text callout box) be scaled automatically when the callout is moved. Use the CustomLength method to specify that the first segment of the callout line retain the fixed length returned by the Length property whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types msoCalloutThree and msoCalloutFour).
- `CustomDrop(Drop As Single)`  
  Sets the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. This distance is measured from the top of the text box unless the AutoAttach property is set to True and the text box is to the left of the origin of the callout line (the place that the callout points to). In this case the drop distance is measured from the bottom of the text box.
    - `Drop As Single` (required): The drop distance, in points.
- `CustomLength(Length As Single)`  
  Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved.
    - `Length As Single` (required): The length of the first segment of the callout, in points.
- `PresetDrop(DropType As MsoCalloutDropType)`  
  Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.
    - `DropType As MsoCalloutDropType` (required): The starting position of the callout line relative to the text bounding box.
