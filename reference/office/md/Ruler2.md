# Ruler2

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C1-0000-0000-C000-000000000046}  

Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.

**Remarks:** Use the Ruler2 property of the TextFrame2 object to return the Ruler2 object that represents the ruler for the text in the specified shape. Use the TabStops2 property to return the TabStops2 object that contains the tab stops on the ruler. Use the Levels property to return the RulerLevels2 object that contains the indentation settings for text outline levels.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes(2).TextFrame2.Ruler2
 .TabStops2.Add ppTabStopLeft, 144
 .Levels(1).FirstMargin = 0
 .Levels(1).LeftMargin = 36
End With
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the Ruler2 object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the Ruler2 object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the Ruler2 object. Read-only.
- `Levels As RulerLevels2  (read-only)`  
  Gets a RulerLevels2 object that represents outline text formatting. Read-only.
- `TabStops As TabStops2  (read-only)`  
  Gets a TabStops2 collection that represents the tab stops for the specified text. Read-only.
