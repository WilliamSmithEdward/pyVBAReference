# NamedSlideShow

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149345C-5A91-11CF-8700-00AA0060263B}  

Represents a custom slide show, which is a named subset of slides in a presentation.

**Remarks:** The NamedSlideShow object is a member of the NamedSlideShows collection. The NamedSlideShows collection contains all the named slide shows in the presentation.

**Example:**

```vba
ActivePresentation.SlideShowSettings _
    .NamedSlideShows("Quick Show").Delete
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Name As String  (read-only)`  
  You cannot use this property to set the name for a custom slide show. Use the Add method to redefine a custom slide show under a new name. Read-only.
- `SlideIDs As Variant  (read-only)`  
  Returns an array of slide IDs for the specified named slide show. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the specified NamedSlideShow object.
