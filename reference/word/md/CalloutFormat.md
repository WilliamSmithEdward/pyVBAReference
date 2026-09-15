# CalloutFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209C5-0000-0000-C000-000000000046}  

Contains properties and methods that apply to line callouts.

**Remarks:** Use the Callout property to return a CalloutFormat object. The following example specifies the following attributes of shape three (a line callout) on the active document: the callout will have a vertical accent bar that separates the text from the callout line; the angle between the callout line and the side of the callout text box will be 30 degrees; there will be no border around the callout text; the callout line will be attached to the top of the callout text box; and the callout line will contain two segments. For this example to work, shape three must be a callout.

## Properties (12)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CalloutFormat object.
- `Accent As MsoTriState  (read/write)`  
  True if a vertical accent bar separates the callout text from the callout line. Read/write MsoTriState.
- `Angle As MsoCalloutAngleType  (read/write)`  
  Returns or sets the angle of the callout line. Read/write MsoCalloutAngleType.
- `AutoLength As MsoTriState  (read-only)`  
  MsoTrue to automatically sets the length of the callout line. Read-only MsoTriState.
- `Border As MsoTriState  (read/write)`  
  Returns or sets whether the text in the specified callout is surrounded by a border. Read/write MsoTriState.
- `Drop As Single  (read-only)`  
  Returns the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. Read-only Single.
- `DropType As MsoCalloutDropType  (read-only)`  
  Returns a value that indicates where the callout line attaches to the callout text box. Read-only MsoCalloutDropType.
- `Gap As Single  (read/write)`  
  Returns or sets the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write Single.
- `Length As Single  (read-only)`  
  Returns the length (in points) of the first segment of the callout line (the segment attached to the text callout box). Read-only Single.
- `Type As MsoCalloutType  (read/write)`  
  Returns or sets the callout type. Read/write MsoCalloutType.

## Methods (3)

- `CustomDrop(Drop As Single)`  
  Sets the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box.
    - `Drop As Single` (required): The drop distance, in points.
- `CustomLength(Length As Single)`  
  Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved.
    - `Length As Single` (required): The length of the first segment of the callout, in points.
- `PresetDrop(DropType As MsoCalloutDropType)`  
  Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.
    - `DropType As MsoCalloutDropType` (required): The starting position of the callout line relative to the text bounding box. If you specify msoCalloutDropCustom, the values of the Drop and AutoAttach properties and the relative positions of the callout text box and callout line origin (the place that the callout points to) are used to determine where the callout line attaches to the text box.
