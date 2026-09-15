# Selection

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493454-5A91-11CF-8700-00AA0060263B}  

Represents the selection in the specified document window. The Selection object is deleted whenever you change slides in an active slide view (the Type property will return ppSelectionNone).

**Example:**

```vba
ActiveWindow.Selection.Copy
```

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Type As PpSelectionType  (read-only)`  
  Represents the type of objects in a selection. Read-only.
- `SlideRange As SlideRange  (read-only)`  
  Returns a SlideRange object that represents a range of selected slides. Read-only.
- `ShapeRange As ShapeRange  (read-only)`  
  Returns a ShapeRange object that represents all the slide objects that have been selected on the specified slide. Read-only.
- `TextRange As TextRange  (read-only)`  
  Returns a TextRange object that represents the selected text. Read-only.
- `ChildShapeRange As ShapeRange  (read-only)`  
  Returns a ShapeRange object that represents the child shapes of a selection.
- `HasChildShapeRange As Boolean  (read-only)`  
  True if the selection contains child shapes. Read-only.
- `TextRange2 As TextRange2  (read-only)`  
  Returns the TextRange2 object of the current Selection object. Read-only.

## Methods (4)

- `Cut()`  
  Deletes the specified object and places it on the Clipboard.
- `Copy()`  
  Copies the specified object to the Clipboard.
- `Delete()`  
  Deletes the specified Selection object.
- `Unselect()`  
  Cancels the current selection.
