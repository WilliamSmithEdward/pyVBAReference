# Sections

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002095A-0000-0000-C000-000000000046}  

A collection of Section objects in a selection, range, or document.

**Remarks:** Use the Sections property to return the Sections collection. The following example inserts text at the end of the last section in the active document. Use the Add method or the InsertBreak method to add a new section to a document. The following example adds a new section at the beginning of the active document. The following example displays the number of sections in the active document, adds a section break above the first paragraph in the selection, and then displays the number of sections again. Use Sections (index), where index is the index number, to return a single Section object. The following example changes the left and right page margins for the first section in the active document.

## Properties (8)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of sections in the collection. Read-only.
- `First As Section  (read-only)`  
  Returns a Section object that represents the first item in the Sections collection.
- `Last As Section  (read-only)`  
  Returns the last item in the Sections collection as a Section object.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Sections object.
- `PageSetup As PageSetup  (read/write)`  
  Returns a PageSetup object that's associated with the specified document, range, section, sections, or selection.

## Methods (2)

- `Item(Index As Long) As Section`  
  Returns an individual Section object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add([Range As Variant], [Start As Variant]) As Section`  
  Returns a Section object that represents a new section added to a document.
    - `Range As Variant` (optional): The range before which you want to insert the section break. If this argument is omitted, the section break is inserted at the end of the document.
    - `Start As Variant` (optional): The type of section break you want to add. Can be one of the WdSectionStart constants. If this argument is omitted, a Next Page section break is added.
