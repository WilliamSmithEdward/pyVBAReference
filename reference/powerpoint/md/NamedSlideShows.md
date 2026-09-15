# NamedSlideShows

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149345B-5A91-11CF-8700-00AA0060263B}  

A collection of all the NamedSlideShow objects in the presentation. Each NamedSlideShow object represents a custom slide show.

**Example:**

```vba
ActivePresentation.SlideShowSettings _
    .NamedSlideShows("Quick Show").Delete
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (2)

- `Item(Index As Variant) As NamedSlideShow`  
  Returns a single NamedSlideShow object from the specified NamedSlideShows collection.
    - `Index As Variant` (required): The name or index number of the single NamedSlideShow object in the collection to be returned.
- `Add(Name As String, safeArrayOfSlideIDs As Variant) As NamedSlideShow`  
  Creates a new named slide show and adds it to the collection of named slide shows in the specified presentation. Returns a NamedSlideShow object that represents the new named slide show.
    - `Name As String` (required): The name of the slide show.
    - `safeArrayOfSlideIDs As Variant` (required): Contains the unique slide IDs of the slides to be displayed in a slide show.
