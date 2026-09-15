# ResampleMediaTasks

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {BA72E554-4FF5-48F4-8215-5505F990966F}  

A collection of ResampleMediaTask objects.

**Remarks:** Use ResampleMediaTasks (index) to return a ResampleMediaTask object, where index is the position of the ResampleMediaTask object to return.

## Properties (2)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `PercentComplete As Long  (read-only)`  
  Returns the percentage of completion of the sample. Read-only.

## Methods (4)

- `Item(Index As Long) As ResampleMediaTask`  
  Returns a single ResampleMediaTask object from the specified ResampleMediaTasks collection.
    - `Index As Long` (required): The index number of the single ResampleMediaTask object in the collection to be returned.
- `Pause()`  
  Pauses the media represented by the specified object.
- `Cancel()`  
  Cancels mobilization of the current media asset.
- `Resume()`  
  Resumes mobilization of the current media asset.
