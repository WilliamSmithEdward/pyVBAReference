# Adjustments

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149347C-5A91-11CF-8700-00AA0060263B}  

Contains a collection of adjustment values for the specified AutoShape, WordArt object, or connector.

**Remarks:** Each adjustment value represents one way an adjustment handle can be adjusted. Because some adjustment handles can be adjusted in two ways for instance, some handles can be adjusted both horizontally and vertically a shape can have more adjustment values than it has adjustment handles. A shape can have up to eight adjustments. Use the Adjustments property to return an Adjustments object. Use Adjustments (_index_), where _index_ is the adjustment value's index number, to return a single adjustment value. Different shapes have different numbers of adjustment values, different kinds of adjustments change the geometry of a shape in different ways, and different kinds of adjustments have different ranges of valid values. For example, the following illustration shows what each of the four adjustment values for a right-arrow callout contributes to the definition of the callout's geometry. !Adjustment values The following table summarizes the ranges of valid adjustment values for different types of adjustments. In most cases, if you specify a value that's beyond the range of valid values, the closest valid value will be assigned to the adjustment.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

Set rac = myDocument.Shapes _
    .AddShape(msoShapeRightArrowCallout, 10, 10, 250, 190)

With rac.Adjustments
    .Item(1) = 0.5    'adjusts width of text box
    .Item(2) = 0.15   'adjusts width of arrow head
    .Item(3) = 0.8    'adjusts length of arrow head
    .Item(4) = 0.4    'adjusts width of arrow neck
End With
```

## Properties (5)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Item As Single  (read/write)`  
  Returns or sets the adjustment value specified by the Index argument. Read/write.
