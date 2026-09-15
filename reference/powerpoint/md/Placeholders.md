# Placeholders

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493476-5A91-11CF-8700-00AA0060263B}  

A collection of all the Shape objects that represent placeholders on the specified slide.

**Remarks:** Each Shape object in the Placeholders collection represents a placeholder for text, a chart, a table, an organizational chart, or some other type of object. If the slide has a title, the title is the first placeholder in the collection. You can delete individual placeholders by using the Deletemethod, and you can restore deleted placeholders by using the AddPlaceholder method, but you cannot add any more placeholders to a slide than it had when it was created. To change the number of placeholders on a given slide, set the Layout property.

**Example:**

```vba
Set sObj = ActivePresentation.Slides.Add(1, ppLayoutText).Shapes
sObj.Title.TextFrame.TextRange.Text = "This is the title text"
sObj.Placeholders(2).TextFrame.TextRange.Text = _
    "Item 1" & Chr(13) & "Item 2"
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (2)

- `Item(Index As Long) As Shape`  
  Returns a single Shape object from the specified Placeholders collection.
    - `Index As Long` (required): The index number of the single Shape object in the collection to be returned.
- `FindByName(Index As Variant) As Shape`  
  Finds the placeholder in the Placeholders collection at the specified index location or with the specified name.
    - `Index As Variant` (required): The index of the placeholder to be found.
