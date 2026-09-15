# TableBackground

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934F8-5A91-11CF-8700-00AA0060263B}  

Represents the background associated with a Table object.

**Remarks:** Use the Background property of a Table object to return the TableBackground object associated with the table. To get a Table object from an existing shape, use the Table property of the Shape or ShapeRange object that contains the table. You can create a shape that contains a table by using the AddTable method of the Shapes collection. The properties of the TableBackground object return objects that represent various aspects of the formatting associated with a table. - Use the Fill property to return a FillFormat object. - Use the Picture property to return a PictureFormat object. - Use the Reflection property to return an ReflectionFormat object. - Use the Shadow property to return a ShadowFormat object.

**Example:**

```vba
Public Sub TableBackground_Example()

    Dim pptShape As PowerPoint.Shape
    Dim pptTable As PowerPoint.Table
    Dim pptTableBackground As PowerPoint.TableBackground
    Dim pptFillFormat As PowerPoint.FillFormat

    Set pptShape = ActivePresentation.Slides(2).Shapes.AddTable(3, 3)
    Set pptTable = pptShape.Table
    Set pptTableBackground = pptTable.Background
    Set pptFillFormat = pptTableBackground.Fill

    ' Add a patterned fill to the table background
    pptFillFormat.Patterned (msoPatternSmallGrid)

    ' Add a shadow to the table background
    pptTableBackground.Shadow.Visible = msoTrue

End Sub
```

## Properties (4)

- `Fill As FillFormat  (read-only)`  
  Returns a FillFormat object that represents the formatting of the fill associated with the table background. Read-only.
- `Picture As PictureFormat  (read-only)`  
  Returns a PictureFormat object that represents the formatting of the picture associated with the table background. Read-only.
- `Reflection As ReflectionFormat  (read-only)`  
  Returns an ReflectionFormat object that represents the reflection effect associated with the table background. Read-only.
- `Shadow As ShadowFormat  (read-only)`  
  Returns a ShadowFormat object that represents the formatting of the shadow associated with the table background. Read-only.
