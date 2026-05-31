# HeaderFooter

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244A1-0000-0000-C000-000000000046}  

Represents a single header or footer. The HeaderFooter object is a member of the HeadersFooters collection.

**Remarks:** You can also return a single HeaderFooter object by using the HeaderFooter property with a Selection object. Use the DifferentFirstPageHeaderFooter property of the PageSetup object to specify a different first page.

**Example:**

```vba
With ActiveSheet.PageSetup
 .CenterHeader = "&D&T"
 .OddAndEvenPagesHeaderFooter = False
 .DifferentFirstPageHeaderFooter = False
 .ScaleWithDocHeaderFooter = True
 .AlignMarginsHeaderFooter = True
End With
```

## Properties (2)

- `Text As String  (read/write)`  
  Returns or sets a Text object that represents text included in the specified header or footer. Read/write.
- `Picture As Graphic  (read-only)`  
  Returns a Picture object that represents a picture field included in the specified header or footer. Read-only.
