# TextStyleLevel

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149349B-5A91-11CF-8700-00AA0060263B}  

Contains character and paragraph formatting information for an outline level.

**Remarks:** The TextStyleLevel object is a member of the TextStyleLevels collection. The TextStyleLevels collection contains one TextStyleLevel object for each of the five outline levels.

**Example:**

```vba
With ActivePresentation.SlideMaster _
        .TextStyles(ppBodyStyle).Levels(1)
    With .Font
        .Name = "Arial"
        .Size = 36
    End With
    With .ParagraphFormat
        .LineRuleBefore = False
        .SpaceBefore = 14
        .Alignment = ppAlignJustify
    End With
End With
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `ParagraphFormat As ParagraphFormat  (read-only)`  
  Returns a ParagraphFormat object that represents paragraph formatting for the specified text. Read-only.
- `Font As Font  (read-only)`  
  Returns a Font object that represents character formatting. Read-only.
