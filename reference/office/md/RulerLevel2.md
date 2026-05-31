# RulerLevel2

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C3-0000-0000-C000-000000000046}  

Contains first-line indent and hanging indent information for an outline level.

**Remarks:** The RulerLevel2 object is a member of the RulerLevels2 collection. The RulerLevels2 collection contains a RulerLevel2 object for each of the five available outline levels.

**Example:**

```vba
With ActivePresentation.SlideMaster _
 .TextStyles(ppBodyStyle).Ruler2.Levels(1)
 .FirstMargin = 9
 .LeftMargin = 54
End With
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the RulerLevel2 object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the RulerLevel2 object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the RulerLevel2 object. Read-only.
- `FirstMargin As Single  (read/write)`  
  Gets or sets the first-line indent for the specified outline level, in points. Read/write.
- `LeftMargin As Single  (read/write)`  
  Gets or sets the left indent for the specified outline level, in points. Read/write.
