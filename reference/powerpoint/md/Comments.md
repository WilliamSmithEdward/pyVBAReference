# Comments

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934D4-5A91-11CF-8700-00AA0060263B}  

Represents a collection of Comment objects.

**Example:**

```vba
Sub CountComments()
    MsgBox "You have " & ActiveWindow.Selection.SlideRange(1) _
        .Comments.Count & " comments on this slide."
End Sub
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (2)

- `Item(Index As Long) As Comment`  
  Returns a single Comment object from the specified Comments collection.
    - `Index As Long` (required): The index number of the single Comment object in the collection to be returned.
- `Add2(Left As Single, Top As Single, Author As String, AuthorInitials As String, Text As String, ProviderID As String, UserID As String) As Comment`  
  Replaces hidden Add method. Returns a Comment object that represents a new comment added to a slide. For more infomation about modern comments, see Modern comments in PowerPoint.
    - `Left As Single` (required): The position, measured in points, of the left edge of the comment, relative to the left edge of the presentation.
    - `Top As Single` (required): The position, measured in points, of the top edge of the comment, relative to the top edge of the presentation.
    - `Author As String` (required): The author of the comment.
    - `AuthorInitials As String` (required): The author's initials.
    - `Text As String` (required): The comment's text.
    - `ProviderID As String` (required): The service that provides contact information.Example: "AD" (Active Directory)
    - `UserID As String` (required): The ID of the user providing the comment.
