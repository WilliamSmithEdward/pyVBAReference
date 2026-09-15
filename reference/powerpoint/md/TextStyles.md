# TextStyles

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493498-5A91-11CF-8700-00AA0060263B}  

A collection of three text styles - title text, body text, and default text - each of which is represented by a TextStyle object.

**Remarks:** Each text style contains a TextFrame object that describes how text is placed within the text bounding box, a Ruler object that contains tab stops and outline indent formatting information, and a TextStyleLevels collection that contains outline text formatting information.

**Example:**

```vba
With ActivePresentation.NotesMaster _
        .TextStyles(ppBodyStyle).TextFrame
    .MarginBottom = 50
    .MarginLeft = 50
    .MarginRight = 50
    .MarginTop = 50
End With
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (1)

- `Item(Type As PpTextStyleType) As TextStyle`  
  Returns a single text style from the specified TextStyles collection.
    - `Type As PpTextStyleType` (required): The text style type.
