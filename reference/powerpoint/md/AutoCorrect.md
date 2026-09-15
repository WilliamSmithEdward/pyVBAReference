# AutoCorrect

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934ED-5A91-11CF-8700-00AA0060263B}  

Represents the AutoCorrect functionality in Microsoft PowerPoint.

**Example:**

```vba
Sub HideAutoCorrectOpButton()

    With Application.AutoCorrect

        .DisplayAutoCorrectOptions = msoFalse

        .DisplayAutoLayoutOptions = msoFalse

    End With

End Sub
```

## Properties (2)

- `DisplayAutoCorrectOptions As Boolean  (read/write)`  
  Determines whether Microsoft PowerPoint should display the AutoCorrect Options button. Read/write.
- `DisplayAutoLayoutOptions As Boolean  (read/write)`  
  Determines whether Microsoft PowerPoint should display the AutoLayout Options button. Read/write.
