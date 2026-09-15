# CaptionLabel

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020979-0000-0000-C000-000000000046}  

Represents a single caption label. The CaptionLabel object is a member of the CaptionLabels collection. The items in the CaptionLabels collection are listed in the Label box in the Caption dialog box.

**Remarks:** Use CaptionLabels (_index_), where _index_ is the caption label name or index number, to return a single CaptionLabel object. The following example sets the numbering style for the Figure caption label. The index number represents the position of the caption label in the CaptionLabels collection. The following example displays the first caption label. Use the Add method to add a custom caption label. The following example adds a caption label named "Photo."

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CaptionLabel object.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `BuiltIn As Boolean  (read-only)`  
  True if the specified caption label is one of the built-in caption labels in Word. Read-only Boolean.
- `ID As WdCaptionLabelID  (read-only)`  
  Returns a WdCaptionLabelID constant that represents the type for the specified caption label if the BuiltIn property of the CaptionLabel object is True. Read-only.
- `IncludeChapterNumber As Boolean  (read/write)`  
  True if a chapter number is included with page numbers or a caption label. Read/write Boolean.
- `NumberStyle As WdCaptionNumberStyle  (read/write)`  
  Returns or sets the number style for the CaptionLabel object. Read/write WdCaptionNumberStyle.
- `ChapterStyleLevel As Long  (read/write)`  
  Returns or sets the heading style that marks a new chapter when chapter numbers are included with the specified caption label. Read/write Long.
- `Separator As WdSeparatorType  (read/write)`  
  Returns or sets the character between the chapter number and the sequence number. Read/write WdSeparatorType.
- `Position As WdCaptionPosition  (read/write)`  
  Returns or sets the position of caption label text. Read/write WdCaptionPosition.

## Methods (1)

- `Delete()`  
  Deletes the specified caption label.
