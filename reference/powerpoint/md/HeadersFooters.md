# HeadersFooters

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493474-5A91-11CF-8700-00AA0060263B}  

Contains all the HeaderFooter objects on the specified slide, notes page, handout, or master.

**Remarks:** Each HeaderFooter object represents a header, footer, date and time, or slide number.

**Example:**

```vba
ActivePresentation.Slides(1).HeadersFooters.Footer _
    .Text = "Volcano Coffee"
```

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `DateAndTime As HeaderFooter  (read-only)`  
  Returns a HeaderFooter object that represents the date and time item that appears in the lower-left corner of a slide or in the upper-right corner of a notes page, handout, or outline. Read-only.
- `SlideNumber As HeaderFooter  (read-only)`  
  Returns a HeaderFooter object that represents the slide number in the lower-right corner of a slide, or the page number in the lower-right corner of a notes page or a page of a printed handout or outline. Read-only.
- `Header As HeaderFooter  (read-only)`  
  Returns a HeaderFooter object that represents the header that appears at the top of a slide or in the upper-left corner of a notes page, handout, or outline. Read-only.
- `Footer As HeaderFooter  (read-only)`  
  Returns a HeaderFooter object that represents the footer that appears at the bottom of a slide or in the lower-left corner of a notes page, handout, or outline. Read-only.
- `DisplayOnTitleSlide As MsoTriState  (read/write)`  
  Determines whether the footer, date and time, and slide number appear on the title slide. Applies to slide masters. Read/write.

## Methods (1)

- `Clear()`  
  Clears the header and footer from the specified slide or slides.
