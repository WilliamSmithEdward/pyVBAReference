# Adjustments

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C0310-0000-0000-C000-000000000046}  

Contains a collection of adjustment values for the specified AutoShape, WordArt object, or connector.

**Remarks:** Each adjustment value represents one way that an adjustment handle can be adjusted. Because some adjustment handles can be adjusted in two ways&mdash;for example, some handles can be adjusted both horizontally and vertically&mdash;a shape can have more adjustment values than it has adjustment handles. A shape can have up to eight adjustments. Use the Adjustments property to return an Adjustments object. Use Adjustments (_index_), where _index_ is the adjustment value's index number, to return a single adjustment value. Different shapes have different numbers of adjustment values. Different kinds of adjustments change the geometry of a shape in different ways. In addition, different kinds of adjustments have different ranges of valid values. For example, the following illustration shows what each of the four adjustment values for a right-arrow callout contributes to the definition of the callout's geometry. !Adjustment The following table summarizes the ranges of valid adjustment values for different types of adjustments. In most cases, if you specify a value that's beyond the range of valid values, the closest valid value will be assigned to the adjustment.

**Example:**

```vba
Set myDocument = Worksheets(1)
Set rac = myDocument.Shapes.AddShape(msoShapeRightArrowCallout, _
 10, 10, 250, 190)
With rac.Adjustments
 .Item(1) = 0.5 'adjusts width of text box
 .Item(2) = 0.15 'adjusts width of arrow head
 .Item(3) = 0.8 'adjusts length of arrow head
 .Item(4) = 0.4 'adjusts width of arrow neck
End With
```

## Properties (5)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns an Integer value that represents the number of objects in the collection.
- `Item As Single  (read/write)`  
  Returns or sets the adjustment value specified by the _Index_ argument. Read/write Single.
