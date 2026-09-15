# TextStyle

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493499-5A91-11CF-8700-00AA0060263B}  

Represents one of three text styles: title text, body text, or default text. Each text style contains a TextFrame object that describes how text is placed within the text bounding box, a Ruler object that contains tab stops and outline indent formatting information, and a TextStyleLevels collection that contains outline text formatting information. The TextStyle object is a member of the TextStyles collection.

**Example:**

```vba
With ActivePresentation.SlideMaster _
        .TextStyles(ppBodyStyle).Levels(1)
    With .Font
        .Name = "Arial"
        .Size = 36
    End With
End With
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Ruler As Ruler  (read-only)`  
  Returns a Ruler object that represents the ruler for the specified text. Read-only.
- `TextFrame As TextFrame  (read-only)`  
  Returns a TextFrame object that contains the alignment and anchoring properties for the specified shape or master text style.
- `Levels As TextStyleLevels  (read-only)`  
  Returns a TextStyleLevels object that represents outline text formatting. Read-only.
