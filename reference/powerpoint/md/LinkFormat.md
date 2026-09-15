# LinkFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493489-5A91-11CF-8700-00AA0060263B}  

Contains properties and methods that apply to linked OLE objects, linked pictures, and IIRC media objects.

**Example:**

```vba
For Each sld In ActivePresentation.Slides

    For Each sh In sld.Shapes

        If sh.Type = msoLinkedOLEObject Then

            If sh.OLEFormat.ProgID = "Excel.Sheet.12" Then

                sh.LinkFormat.AutoUpdate = ppUpdateOptionManual

            End If

        End If

    Next

Next
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `SourceFullName As String  (read/write)`  
  Returns or sets the name and path of the source file for the linked OLE object. Read/write.
- `AutoUpdate As PpUpdateOption  (read/write)`  
  Returns or sets the way the link will be updated. Read/write.

## Methods (2)

- `Update()`  
  Updates the specified linked OLE object.
- `BreakLink()`  
  Breaks the link between the source file and the specified OLE object, picture, or linked field.
