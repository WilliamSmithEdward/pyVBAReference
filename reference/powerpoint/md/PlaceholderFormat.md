# PlaceholderFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493477-5A91-11CF-8700-00AA0060263B}  

Contains properties that apply specifically to placeholders, such as placeholder type.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes.Placeholders
    If .Count > 0 Then
        With .Item(1)
            Select Case .PlaceholderFormat.Type
                Case ppPlaceholderTitle
                    .TextFrame.TextRange = "Title Text"

                Case ppPlaceholderCenterTitle
                    .TextFrame.TextRange = "Centered Title Text"

                Case Else
                    MsgBox "There's no horizontal " _
                        "title on this slide"
            End Select
        End With
    End If
End With
```

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Type As PpPlaceholderType  (read-only)`  
  Represents the type of placeholder. Read-only.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write.
- `ContainedType As MsoShapeType  (read-only)`  
  Returns the type of the shape that contains the specified placeholder format. Read-only.
- `Position As Long  (read-only)`
