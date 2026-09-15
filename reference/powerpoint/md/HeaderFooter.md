# HeaderFooter

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149349C-5A91-11CF-8700-00AA0060263B}  

Represents a header, footer, date and time, slide number, or page number on a slide or master. All the HeaderFooter objects for a slide or master are contained in a HeadersFooters object.

**Remarks:** Use one of the properties listed in the following table to return the HeaderFooter object.

**Example:**

```vba
ActivePresentation.Slides(1).HeadersFooters.Footer.Text = "Volcano Coffee"
```

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.
- `Text As String  (read/write)`  
  Returns or sets a String that represents the text contained in the specified object. Read/write.
- `UseFormat As MsoTriState  (read/write)`  
  Determines whether the date and time object contains automatically updated information. Read/write.
- `Format As PpDateTimeFormat  (read/write)`  
  Returns or sets the format for the automatically updated date and time. Read/write.
