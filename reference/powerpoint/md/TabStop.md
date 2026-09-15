# TabStop

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493494-5A91-11CF-8700-00AA0060263B}  

Represents a single tab stop. The TabStop object is a member of the TabStops collection. The TabStops collection represents all the tab stops on one ruler.

**Example:**

```vba
ActivePresentation.Slides(1).Shapes(2).TextFrame _
    .Ruler.TabStops(1).Clear
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Type As PpTabStopType  (read/write)`  
  Represents the formatting of a tab stop. Read/write.
- `Position As Single  (read/write)`  
  Returns or sets the position of the specified tab stop, in points. Read/write.

## Methods (1)

- `Clear()`  
  Clears the specified tab stop from the ruler and deletes it from the TabStops collection.
