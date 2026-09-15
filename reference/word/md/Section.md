# Section

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020959-0000-0000-C000-000000000046}  

Represents a single section in a selection, range, or document. The Section object is a member of the Sections collection. The Sections collection includes all the sections in a selection, range, or document.

**Remarks:** Use Sections (Index), where Index is the index number, to return a single Section object. The following example changes the left and right page margins for the first section in the active document. Use the Add method or the InsertBreak method to add a new section to a document. The following example adds a new section at the beginning of the active document. The following example adds a section break above the first paragraph in the selection.

## Properties (10)

- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that's contained in the specified object.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Section object.
- `PageSetup As PageSetup  (read/write)`  
  Returns a PageSetup object that is associated with the specified section.
- `Headers As HeadersFooters  (read-only)`  
  Returns a HeadersFooters collection that represents the headers for the specified section. Read-only.
- `Footers As HeadersFooters  (read-only)`  
  Returns a HeadersFooters collection that represents the footers in the specified section. Read-only.
- `ProtectedForForms As Boolean  (read/write)`  
  True if the specified section is protected for forms. Read/write Boolean.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders in the section.
