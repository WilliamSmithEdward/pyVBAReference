# GroupShapes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209B6-0000-0000-C000-000000000046}  

Represents the individual shapes within a grouped shape. Each shape contained within a group of shapes is represented by a Shape object.

**Remarks:** Use the GroupItems property to return the GroupShapes collection. Use GroupItems (Index), where Index is the number of the individual shape within the grouped shape, to return a single shape from the GroupShapes collection. The following example adds three triangles to the active document, groups them, sets a color for the entire group, and then changes the color for the second triangle only.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified GroupShapes object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of shapes in the collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Item(Index As Variant) As Shape`  
  Returns an individual Shape object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Range(Index As Variant) As ShapeRange`  
  Returns a ShapeRange object.
    - `Index As Variant` (required): Specifies which shapes are to be included in the specified range. Can be an integer that specifies the index number of a shape within the Shapes collection, a string that specifies the name of a shape, or a array that contains integers or strings.
