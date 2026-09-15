# Coauthoring

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {BA72E557-4FF5-48F4-8215-5505F990966F}  

Contains properties and methods for manipulating co authoring in a presentation.

**Remarks:** The Presentation.Coauthoring property returns a Coauthoring object.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `MergeMode As Boolean  (read-only)`  
  Returns True if the application is in merge mode. Read-only.
- `FavorServerEditsDuringMerge As Boolean  (read/write)`  
  Gets or sets whether the merged document favors server-side edits when conflicts occur. Read/write.
- `PendingUpdates As Boolean  (read-only)`  
  Returns True if merge-mode updates are pending. Read-only.

## Methods (1)

- `EndReview()`  
  Terminates merge mode and ends the review.
