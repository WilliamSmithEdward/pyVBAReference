# Ruler

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493490-5A91-11CF-8700-00AA0060263B}  

Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes(2).TextFrame.Ruler

    .TabStops.Add ppTabStopLeft, 144

    .Levels(1).FirstMargin = 0

    .Levels(1).LeftMargin = 36

End With
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `TabStops As TabStops  (read-only)`  
  Returns a TabStops collection that represents the tab stops for the specified text. Read-only.
- `Levels As RulerLevels  (read-only)`  
  Returns a RulerLevels object that represents outline indent formatting. Read-only.
