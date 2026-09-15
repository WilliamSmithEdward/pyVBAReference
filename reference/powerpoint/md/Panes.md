# Panes

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934CB-5A91-11CF-8700-00AA0060263B}  

A collection of Pane objects that represent the slide, outline, and notes panes in the document window for normal view, or the single pane of any other view in the document window.

**Remarks:** In normal view, the Panes collection contains three members. All other document window views have only a single pane, resulting in a Panes collection with one member.

**Example:**

```vba
With ActiveWindow

    If .Panes.Count = 1 Then

        .ViewType = ppViewNormal

        .SplitHorizontal = 15

    End If

End With
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (1)

- `Item(Index As Long) As Pane`  
  Returns a single Pane object from the specified Panes collection.
    - `Index As Long` (required): The index number of the single Pane object in the collection to be returned.
