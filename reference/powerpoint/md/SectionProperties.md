# SectionProperties

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {BA72E551-4FF5-48F4-8215-5505F990966F}  

Contains properties and methods for manipulating sections in a presentation.

**Remarks:** Use the SectionProperties property of the Presentation object to get a SectionProperties object.

## Properties (3)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.

## Methods (9)

- `Name(sectionIndex As Long) As String`  
  Returns the name of the specified section as a String.
    - `sectionIndex As Long` (required): The index of the section in the SectionProperties collection.
- `Rename(sectionIndex As Long, sectionName As String)`  
  Renames the specified section with the specified name.
    - `sectionIndex As Long` (required): The index of the section to rename.
    - `sectionName As String` (required): The new name of the section.
- `SlidesCount(sectionIndex As Long) As Long`  
  Returns the number of slides in the specified section.
    - `sectionIndex As Long` (required): The index of the section.
- `FirstSlide(sectionIndex As Long) As Long`  
  Returns the index of the first slide in the specified section, or returns -1 if the section is empty.
    - `sectionIndex As Long` (required): The index of the section.
- `AddBeforeSlide(SlideIndex As Long, sectionName As String) As Long`  
  Adds a section immediately before the specified slide index, and returns the index of the new section.
    - `SlideIndex As Long` (required): The index of the slide before which to add the section.
    - `sectionName As String` (required): The name of the new section.
- `AddSection(sectionIndex As Long, [sectionName As Variant]) As Long`  
  Adds a new section at the specified index position and returns the index of the newly created section.
    - `sectionIndex As Long` (required): The index of the section before which to add the section.
    - `sectionName As Variant` (optional): The name of the new section.
- `Move(sectionIndex As Long, toPos As Long)`  
  Moves the specified section to the specified index position, moving the slides in the section along with the section break.
    - `sectionIndex As Long` (required): The index of the section to move.
    - `toPos As Long` (required): The index position to which to move the section.
- `Delete(sectionIndex As Long, deleteSlides As Boolean)`  
  Deletes the section break that sets off the specified section, and optionally deletes all the slides in the section.
    - `sectionIndex As Long` (required): The index of the section to delete.
    - `deleteSlides As Boolean` (required): Whether to delete all the slides in the section. True, to delete all the slides within the section; False not to delete them.
- `SectionID(sectionIndex As Long) As String`  
  Returns a string that represents the unique identifier of the specified section.
    - `sectionIndex As Long` (required): The index of the section.
