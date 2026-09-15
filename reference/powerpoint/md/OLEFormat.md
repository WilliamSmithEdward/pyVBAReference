# OLEFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493488-5A91-11CF-8700-00AA0060263B}  

Contains properties and methods that apply to OLE objects.

**Remarks:** The LinkFormat object contains properties and methods that apply to linked OLE objects only. The PictureFormat object contains properties and methods that apply to pictures and OLE objects.

**Example:**

```vba
For Each sld In ActivePresentation.Slides

    For Each sh In sld.Shapes

        If sh.Type = msoLinkedOLEObject Then

            If sh.OLEFormat.ProgID = "Excel.Sheet" Then

                sh.LinkFormat.AutoUpdate = ppUpdateOptionManual

            End If

        End If

    Next

Next
```

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `ObjectVerbs As ObjectVerbs  (read-only)`  
  Returns a ObjectVerbs collection that contains all the OLE verbs for the specified OLE object. Read-only.
- `Object As Object  (read-only)`  
  Returns the object that represents the specified OLE object's top-level interface. Read-only.
- `ProgID As String  (read-only)`  
  Returns the programmatic identifier (ProgID) for the specified OLE object. Read-only.
- `FollowColors As PpFollowColors  (read/write)`  
  Returns or sets the extent to which the colors in the specified object follow the slide's color scheme. Read/write.

## Methods (2)

- `DoVerb([Index As Long])`  
  Requests that an OLE object perform one of its verbs.
    - `Index As Long` (optional): The verb to perform. If this argument is omitted, the default verb is performed.
- `Activate()`  
  Activates the specified object.
