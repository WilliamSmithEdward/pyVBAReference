# PrintOptions

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149345D-5A91-11CF-8700-00AA0060263B}  

Contains print options for a presentation.

**Example:**

```vba
With ActivePresentation
    With .PrintOptions
        .NumberOfCopies = 2
        .Collate = False
        .PrintColorType = ppPrintColor
        .PrintHiddenSlides = True
        .FitToPage = True
        .FrameSlides = True
        .OutputType = ppPrintOutputSlides
    End With
    .PrintOut
End With
```

## Properties (19)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `PrintColorType As PpPrintColorType  (read/write)`  
  Returns or sets the way the specified document will be printed: in black and white, in pure black and white (also referred to as high contrast), or in color. Read/write.
- `Collate As MsoTriState  (read/write)`  
  Determines whether a complete copy of the specified presentation is printed before the first page of the next copy is printed. Read/write.
- `FitToPage As MsoTriState  (read/write)`  
  Determines whether the slides will be scaled to fill the page they're printed on. Read/write.
- `FrameSlides As MsoTriState  (read/write)`  
  Determines whether a thin frame is placed around the border of the printed slides. Read/write.
- `NumberOfCopies As Long  (read/write)`  
  Returns or sets the number of copies of a presentation to be printed. Read/write.
- `OutputType As PpPrintOutputType  (read/write)`  
  Returns or sets a value that indicates which component (slides, handouts, notes pages, or an outline) of the presentation is to be printed. Read/write.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `PrintHiddenSlides As MsoTriState  (read/write)`  
  Determines whether hidden slides in the specified presentation will be printed. Read/write.
- `PrintInBackground As MsoTriState  (read/write)`  
  Determines whether the specified presentation is printed in the background. Read/write.
- `RangeType As PpPrintRangeType  (read/write)`  
  Returns or sets the type of print range for the presentation. Read/write.
- `Ranges As PrintRanges  (read-only)`  
  Returns the PrintRanges object, which represents the ranges of slides in the presentation to be printed. Read-only.
- `PrintFontsAsGraphics As MsoTriState  (read/write)`  
  Determines whether TrueType fonts are printed as graphics. Read/write.
- `SlideShowName As String  (read/write)`  
  Returns or sets the name of the custom slide show to print. Read/write .
- `ActivePrinter As String  (read/write)`  
  Returns the name of the active printer. Read-only.
- `HandoutOrder As PpPrintHandoutOrder  (read/write)`  
  Returns or sets the page layout order in which slides appear on printed handouts that show multiple slides on one page. Read/write.
- `PrintComments As MsoTriState  (read/write)`  
  Sets or returns whether comments will be printed. Read/write.
- `sectionIndex As Long  (read/write)`  
  Returns the index of the selected section in the PrintOptions object. Read/write.
- `HighQuality As MsoTriState  (read/write)`  
  Indicates whether to print in high quality. Read/write.
