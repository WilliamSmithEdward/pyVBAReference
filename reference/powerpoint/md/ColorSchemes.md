# ColorSchemes

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149346E-5A91-11CF-8700-00AA0060263B}  

A collection of all the ColorScheme objects in the specified presentation. Each ColorScheme object represents a color scheme, which is a set of colors that are used together on a slide.

**Example:**

```vba
ActivePresentation.ColorSchemes(2).Delete
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (2)

- `Item(Index As Long) As ColorScheme`  
  Returns a single ColorScheme object from the specified ColorSchemes collection.
    - `Index As Long` (required): The index number of the single ColorScheme object in the collection to be returned.
- `Add([Scheme As ColorScheme]) As ColorScheme`  
  Adds a color scheme to the collection of available schemes. Returns a ColorScheme object that represents the added color scheme.
    - `Scheme As ColorScheme` (optional): The color scheme to add. Can be a ColorScheme object from any slide or master or an item in the ColorSchemes collection from any open presentation. If you omit this parameter, the first ColorScheme object (the first standard color scheme) in the specified presentation's ColorSchemes collection is used.
