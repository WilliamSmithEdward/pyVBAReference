# Adjustments

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209C4-0000-0000-C000-000000000046}  

Contains a collection of adjustment values for the specified AutoShape or WordArt object. Each adjustment value represents one way an adjustment handle can be adjusted. Because some adjustment handles can be adjusted in two ways&mdash;for instance, some handles can be adjusted both horizontally and vertically&mdash;a shape can have more adjustment values than it has adjustment handles. A shape can have up to eight adjustments.

**Remarks:** Use the Adjustments property to return an Adjustments object. Use Adjustments (index), where index is the adjustment value's index number, to return a single adjustment value. Different shapes have different numbers of adjustment values, different kinds of adjustments change the geometry of a shape in different ways, and different kinds of adjustments have different ranges of valid values. The following table summarizes the ranges of valid adjustment values for different types of adjustments. In most cases, if you specify a value that's beyond the range of valid values, the closest valid value will be assigned to the adjustment. The following example adds a right-arrow callout to the active document and sets adjustment values for the callout. Note that although the shape has only three adjustment handles, it has four adjustments. Adjustments three and four both correspond to the handle between the head and neck of the arrow.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified object. This is usually a Shape or ShapeRange object.
- `Count As Long  (read-only)`  
  Returns the number of items in the Adjustments collection. Read-only Long.
- `Item As Single  (read/write)`  
  Returns or sets the adjustment value specified by the Index argument. Read/write Single.
