# Slides

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493469-5A91-11CF-8700-00AA0060263B}  

A collection of all the Slide objects in the specified presentation.

**Remarks:** If your Visual Studio solution includes the Microsoft.Office.Interop.PowerPoint reference, this collection maps to the following types: - Microsoft.Office.Interop.PowerPoint.Slides.GetEnumerator (to enumerate the Slide objects.) The following examples describe how to: - Create a slide and add it to the collection - Return a single slide that you specify by name, index number, or slide ID number - Return a subset of the slides in the presentation - Apply a property or method to all the slides in the presentation at the same time

**Example:**

```vba
ActivePresentation.Slides.Add 2, ppLayoutBlank
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (6)

- `Item(Index As Variant) As Slide`  
  Returns a single Slide object from the specified Slides collection.
    - `Index As Variant` (required): The name or index number of the single Slide object in the collection to be returned.
- `FindBySlideID(SlideID As Long) As Slide`  
  Returns a Slide object that represents the slide with the specified slide ID number. Each slide is automatically assigned a unique slide ID number when it is created. Use the SlideID property to return a slide's ID number.
    - `SlideID As Long` (required): Specifies the ID number of the slide you want to return. Microsoft PowerPoint assigns this number when the slide is created.
- `InsertFromFile(FileName As String, Index As Long, [SlideStart As Long], [SlideEnd As Long]) As Long`  
  Inserts slides from a file into a presentation, at the specified location. Returns an Integer that represents the number of slides inserted.
    - `FileName As String` (required): The name of the file that contains the slides you want to insert.
    - `Index As Long` (required): The index number of the Slide object in the specified Slides collection you want to insert the new slides after.
    - `SlideStart As Long` (optional): The index number of the first Slide object in the Slides collection in the file denoted by FileName.
    - `SlideEnd As Long` (optional): The index number of the last Slide object in the Slides collection in the file denoted by FileName.
- `Range([Index As Variant]) As SlideRange`  
  Returns a SlideRange object that represents a subset of the slides in a Slides collection.
    - `Index As Variant` (optional): The individual slides that are to be included in the range. Can be an Integer that specifies the index number of the slide, a String that specifies the name of the slide, or an array that contains either integers or strings. If this argument is omitted, the Range method returns all the objects in the specified collection.
- `Paste([Index As Long]) As SlideRange`  
  Pastes the slides on the Clipboard into the Slides collection for the presentation. Specify where you want to insert the slides with the Index argument. Returns a SlideRange object that represents the pasted objects. Each pasted slide becomes a member of the specified Slides collection.
    - `Index As Long` (optional): The index number of the slide that the slides on the Clipboard are to be pasted before. If this argument is omitted, the slides on the Clipboard are pasted after the last slide in the presentation.
- `AddSlide(Index As Long, pCustomLayout As CustomLayout) As Slide`  
  Creates a new slide, adds it to the Slides collection, and returns the slide.
    - `Index As Long` (required): The index of the slide to be added.
    - `pCustomLayout As CustomLayout` (required): The layout of the slide.
