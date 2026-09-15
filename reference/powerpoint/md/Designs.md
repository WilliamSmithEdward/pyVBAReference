# Designs

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934D6-5A91-11CF-8700-00AA0060263B}  

Represents a collection of slide design templates.

**Remarks:** Use the Designs property of the Presentation object to reference a design template. To add or clone an individual design template, use the Designs collection's Add or Clone methods, respectively. To refer to an individual design template, use the Item method. To load a design template, use the Load method.

**Example:**

```vba
Sub AddDesignMaster()

    With ActivePresentation.Designs

        .Add designName:="MyDesignName"

        MsgBox .Item("MyDesignName").Name

    End With

End Sub
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (4)

- `Item(Index As Variant) As Design`  
  Returns a single Design object from the specified Designs collection.
    - `Index As Variant` (required): The name or index number of the single Design object in the collection to be returned.
- `Add(designName As String, [Index As Long]) As Design`  
  Returns a Design object that represents a new slide design.
    - `designName As String` (required): The name of the design.
    - `Index As Long` (optional): The index number of the design in the Designs collection. The default value is -1, which means that if you omit the Index parameter, the new slide design is added at the end of existing slide designs.
- `Load(TemplateName As String, [Index As Long]) As Design`  
  Returns a Design object that represents a design loaded into the master list of the specified presentation.
    - `TemplateName As String` (required): The path to the design template.
    - `Index As Long` (optional): The index number of the design template in the collection of design templates. The default is -1, which means the design template is added to the end of the list of designs in the presentation.
- `Clone(pOriginal As Design, [Index As Long]) As Design`  
  Creates a copy of a Design object.
    - `pOriginal As Design` (required): Design object. The original design.
    - `Index As Long` (optional): The index location in the Designs collection into which the design will be copied. If Index is omitted, the cloned design is added to the end of the Designs collection.
