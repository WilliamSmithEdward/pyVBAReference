# StoryRanges

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002098C-0000-0000-C000-000000000046}  

A collection of Range objects that represent stories in a document.

**Remarks:** Use the StoryRanges property to return the StoryRanges collection. The following example removes manual character formatting from the text in all stories other than the main text story in the active document. The Add method is not available for the StoryRanges collection. The number of stories in the StoryRanges collection is finite. Use StoryRanges (Index), where Index is a WdStoryType constant, to return a single story as a Range object. The following example adds text to the primary header story and then displays the text. The following example copies the text of the footnotes from the active document into a new document. If you attempt to return a story that is not available in the specified document, an error occurs. The following example determines whether a footnote story is available in the active document. Use the NextStoryRange property to loop through all stories in a document. The following example searches each story in the active document for the text "Microsoft Word." When the text is found, it is formatted as italic.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of story ranges in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified StoryRanges object.

## Methods (1)

- `Item(Index As WdStoryType) As Range`  
  Returns a single story of a range or selection as a Range object.
    - `Index As WdStoryType` (required): The specified story type.
