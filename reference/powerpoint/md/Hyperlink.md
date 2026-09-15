# Hyperlink

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493465-5A91-11CF-8700-00AA0060263B}  

Represents a hyperlink associated with a non-placeholder shape or text.

**Remarks:** Use a hyperlink to jump to an Internet or intranet site, to another file, or to a slide within the active presentation. The Hyperlink object is a member of the Hyperlinks collection. The Hyperlinks collection contains all the hyperlinks on a slide or a master.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes(3) _

        .ActionSettings(ppMouseClick)

    .Action = ppActionHyperlink

    .Hyperlink.Address = "https://www.microsoft.com"

End With
```

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Type As MsoHyperlinkType  (read-only)`  
  Represents the type of hyperlink. Read-only.
- `Address As String  (read/write)`  
  Returns or sets the Internet address (URL) to the target document. Read/write.
- `SubAddress As String  (read/write)`  
  Returns or sets the location within a document - such as a bookmark in a word document, a range in a Microsoft Office Excel worksheet, or a slide in a Microsoft PowerPoint presentation - associated with the specified hyperlink. Read/write.
- `EmailSubject As String  (read/write)`  
  Returns or sets the text string of the hyperlink subject line. The subject line is appended to the Internet address (URL) of the hyperlink. Read/write.
- `ScreenTip As String  (read/write)`  
  Returns or sets the ScreenTip text of a hyperlink. Read/write.
- `TextToDisplay As String  (read/write)`  
  Returns or sets the display text for a hyperlink not associated with a graphic. Read/write.
- `ShowAndReturn As MsoTriState  (read/write)`  
  Determines if and under what circumstances Microsoft PowerPoint returns to the initiating slide show. Read/write.

## Methods (4)

- `AddToFavorites()`  
  Adds a shortcut that represents the specified hyperlink's target document to the Favorites folder in the Windows folder.
- `Follow()`  
  Displays the HTML document associated with the specified hyperlink in a new web browser window.
- `CreateNewDocument(FileName As String, EditNow As MsoTriState, Overwrite As MsoTriState)`  
  Creates a new Web presentation associated with the specified hyperlink.
    - `FileName As String` (required): The path and file name of the document.
    - `EditNow As MsoTriState` (required): Determines whether the document is opened immediately in its associated editor.
    - `Overwrite As MsoTriState` (required): Determines whether any existing file of the same name in the same folder is overwritten.
- `Delete()`  
  Deletes the specified Hyperlink object.
