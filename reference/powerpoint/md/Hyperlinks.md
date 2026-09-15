# Hyperlinks

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493464-5A91-11CF-8700-00AA0060263B}  

A collection of all the Hyperlink objects on a slide or master.

**Example:**

```vba
For Each hl In ActivePresentation.Slides(1).Hyperlinks

    If hl.Address = "c:\current work\sales.ppt" Then

        hl.Address = "c:\new\newsales.ppt"

    End If

Next
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (1)

- `Item(Index As Long) As Hyperlink`  
  Returns a single Hyperlink object from the specified Hyperlinks collection.
    - `Index As Long` (required): The index number of the single Hyperlink object in the collection to be returned.
