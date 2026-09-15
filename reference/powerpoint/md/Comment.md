# Comment

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934D5-5A91-11CF-8700-00AA0060263B}  

Represents a comment on a given slide or slide range. The Comment object is a member of the Comments collection object.

**Remarks:** Use the following properties to access comment data:

**Example:**

```vba
Sub ShowComment()

    With ActivePresentation.Slides(1).Comments

        If .Count > 0 Then

            MsgBox "The first comment on this slide is by " & .Item(1).Author

        Else

            MsgBox "There are no comments on this slide."

        End If

    End With

End Sub
```

## Properties (14)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Author As String  (read-only)`  
  Returns a String that represents the author as for a specified Comment object. Read-only.
- `AuthorInitials As String  (read-only)`  
  Returns the author's initials as a read-only String for a specified Comment object. Read-only.
- `Text As String  (read-only)`  
  Returns a String that represents the text in a comment. Read-only.
- `DateTime As Date  (read-only)`  
  Returns the date and time a comment was created.
- `AuthorIndex As Long  (read-only)`  
  Returns a Long representing the index number of a comment for a given author. The first comment for a given author has an index number of 1, their second comment has an index number of 2. Read-only.
- `Left As Single  (read-only)`  
  Returns a Single that represents the distance in points from the left edge of the comment to the left edge of the slide. Read-only.
- `Top As Single  (read-only)`  
  Returns a Single that represents the distance in points from the left edge of the comment to the left edge of the slide. Read-only.
- `ProviderID As String  (read-only)`  
  Returns a String that represents ID of the service providing contact information. Read-only.
- `UserID As String  (read-only)`  
  Returns a String that represents user ID of the contact. Read-only.
- `TimeZoneBias As Long  (read-only)`  
  Returns a String that specifies the local time zone adjustment for the contact. Read-only.
- `Replies As Comments  (read-only)`  
  Returns a Comments collection of Comment objects that are children of the specified comment. Read-only.
- `Collapsed As Boolean  (read-only)`  
  Returns whether the replies to a comment are shown (expanded) or hidden (collapsed). Ready-only.

## Methods (1)

- `Delete()`  
  Deletes the specified Comment object.
