# TextRetrievalMode

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020939-0000-0000-C000-000000000046}  

Represents options that control how text is retrieved from a Range object.

**Remarks:** Use the TextRetrievalMode property to return a TextRetrievalMode object. The following example displays the text of the first sentence in the active document, excluding field codes and hidden text. Changing the ViewType, IncludeHiddentText, or IncludeFieldCodes property of the TextRetrievalMode object doesn't change the screen display. Instead, changing one of these properties determines what text is retrieved from a Range object when the Text property is used.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TextRetrievalMode object.
- `ViewType As WdViewType  (read/write)`  
  Returns or sets the view for the TextRetrievalMode object. Read/write WdViewType.
- `Duplicate As TextRetrievalMode  (read-only)`  
  Returns a read-only TextRetrievalMode object that represents options related to retrieving text from a Range object.
- `IncludeHiddenText As Boolean  (read/write)`  
  True if the text retrieved from the specified range includes hidden text. Read/write Boolean.
- `IncludeFieldCodes As Boolean  (read/write)`  
  True if the text retrieved from the specified range includes field codes. Read/write Boolean.
