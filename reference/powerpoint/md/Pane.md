# Pane

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934CC-5A91-11CF-8700-00AA0060263B}  

An object representing one of the three panes in normal view or the single pane of any other view in the document window.

**Remarks:** Use Panes (_index_), where _index_ is the index number for a pane, to return a single Pane object. The following table lists the names of the panes in normal view with their corresponding index numbers. When using a document window view other than normal view, use Panes (1) to reference the single Pane object. Use the Activatemethod to make the specified pane active. Use the ViewTypeproperty to determine which pane is active. Normal view is the only view with multiple panes. All other document window views have only a single pane, which is the document window.

**Example:**

```vba
With ActiveWindow

    If .ActivePane.ViewType = ppViewSlide Then

        .Panes(3).Activate

    End If

End With
```

## Properties (4)

- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Active As MsoTriState  (read-only)`  
  Returns whether the specified pane or window is active. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `ViewType As PpViewType  (read-only)`  
  Returns the type of view for the specified pane. Read-only.

## Methods (1)

- `Activate()`  
  Activates the specified object.
