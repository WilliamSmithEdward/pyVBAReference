# Tags

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934B9-5A91-11CF-8700-00AA0060263B}  

Represents a tag or a custom property that you can create for a shape, slide, or presentation.

**Remarks:** Each Tags object contains the name of a custom property and a value for that property. Create tags when you want to be able to selectively work with specific members of a collection, based on an attribute that isn't already represented by a built-in property. For example, if you want to be able to categorize slides in a presentation based on what region of the country/region they apply to, you could create a Region tag and assign a Region value to each slide in the presentation. You could then selectively perform an operation on some of the slides, based on the values of their Region tags, such as hiding all the slides with the Region value "East."

**Example:**

```vba
ActivePresentation.Slides(1).Tags.Add "Region", "East"
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (5)

- `Item(Name As String) As String`  
  Returns a single tag from the specified Tags collection.
    - `Name As String` (required): The name of the single tag in the collection to be returned.
- `Add(Name As String, Value As String)`  
  Adds a tag to the Tags collection of an object.
    - `Name As String` (required): The name of the tag.
    - `Value As String` (required): The value of the tag.
- `Delete(Name As String)`  
  Deletes a tag.
    - `Name As String` (required): Tthe name of the tag to be deleted.
- `Name(Index As Long) As String`  
  Returns the name of the specified tag as a String.
    - `Index As Long` (required): The tag number.
- `Value(Index As Long) As String`  
  Returns the value of the specified tag as a String.
    - `Index As Long` (required): The tag number.
