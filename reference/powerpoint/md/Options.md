# Options

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934EE-5A91-11CF-8700-00AA0060263B}  

Represents application options in Microsoft PowerPoint.

**Example:**

```vba
Sub TogglePasteOptionsButton()

    With Application.Options

        If .DisplayPasteOptions = False Then

            .DisplayPasteOptions = True

        End If

    End With

End Sub
```

## Properties (1)

- `DisplayPasteOptions As MsoTriState  (read/write)`  
  Determines whether Microsoft PowerPoint displays the Paste Options button, which appears directly under newly pasted text. Read/write.
