# ColorScheme

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149346F-5A91-11CF-8700-00AA0060263B}  

Represents a color scheme, which is a set of eight colors used for the different elements of a slide, notes page, or handout, such as the title or background. (Note that the color schemes for slides, notes pages, and handouts in a presentation can be set independently.)

**Remarks:** Each color is represented by an RGBColor object. The ColorScheme object is a member of the ColorSchemes collection. The ColorSchemes collection contains all the color schemes in a presentation. The following examples describe how to do the following: - Return a ColorScheme object from the collection of all the color schemes in the presentation - Return the ColorScheme object attached to a specific slide or master - Return the color of a single slide element from a ColorScheme object

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

- `Colors(SchemeColor As PpColorSchemeIndex) As RGBColor`  
  Returns an RGBColor object that represents a single color in a color scheme.
    - `SchemeColor As PpColorSchemeIndex` (required): The individual color in the specified color scheme.
- `Delete()`  
  Deletes the specified ColorScheme object.
