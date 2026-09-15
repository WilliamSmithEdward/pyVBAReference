# RevisionsFilter

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {D523C26B-7278-4FA9-AA0B-0827DC8B41CE}  

Represents the current settings related to display of reviewers' comments and revision marks in the document.

**Remarks:** Use the View.RevisionsFilter property to return a RevisionsFilter object.

## Properties (3)

- `View As WdRevisionsView  (read/write)`  
  Sets or returns a WdRevisionsView constant that represents the global option that specifies whether Word displays the original version of a document or the final version, which might have revisions and formatting changes applied. Read/write.
- `Markup As WdRevisionsMarkup  (read/write)`  
  Returns or sets a WdRevisionsMarkup constant that specifies the extent of reviewer markup displayed in the document. Read/write.
- `Reviewers As Reviewers  (read-only)`  
  Returns a Reviewers object that represents the collection of reviewers of one or more documents.

## Methods (1)

- `ToggleShowAllReviewers()`  
  Shows or hides all revisions in a document that contains comments and tracked changes.
