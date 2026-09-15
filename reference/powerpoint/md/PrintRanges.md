# PrintRanges

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149345E-5A91-11CF-8700-00AA0060263B}  

A collection of all the PrintRange objects in the specified presentation. Each PrintRange object represents a range of consecutive slides or pages to be printed.

**Example:**

```vba
ActivePresentation.PrintOptions.Ranges.ClearAll
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (3)

- `Add(Start As Long, End As Long) As PrintRange`  
  Returns a PrintRange object that represents a new print range to be added to the PrintRanges collection.
    - `Start As Long` (required): The number of the slide that is the starting slide in the range.
    - `End As Long` (required): The number of the slide that is the ending slide in the range.
- `ClearAll()`  
  Clears all the print ranges from the PrintRanges collection. Use the Add method of the PrintRanges collection to add print ranges to the collection.
- `Item(Index As Long) As PrintRange`  
  Returns a single PrintRange object from the specified PrintRanges collection.
    - `Index As Long` (required): The index number of the single PrintRange object in the collection to be returned.
