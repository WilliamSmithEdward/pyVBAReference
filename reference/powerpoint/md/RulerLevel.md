# RulerLevel

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493492-5A91-11CF-8700-00AA0060263B}  

Contains first-line indent and hanging indent information for an outline level.

**Remarks:** The RulerLevel object is a member of the RulerLevels collection. The RulerLevels collection contains a RulerLevel object for each of the five available outline levels.

**Example:**

```vba
With ActivePresentation.SlideMaster _
        .TextStyles(ppBodyStyle).Ruler.Levels(1)
    .FirstMargin = 9
    .LeftMargin = 54
End With
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `FirstMargin As Single  (read/write)`  
  Returns or sets the first-line indent for the specified outline level, in points. Read/write.
- `LeftMargin As Single  (read/write)`  
  Returns or sets the left indent for the specified outline level, in points. Read/write.
