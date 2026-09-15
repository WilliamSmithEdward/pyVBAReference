# TableOfAuthorities

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020911-0000-0000-C000-000000000046}  

Represents a single table of authorities in a document (a TOA field). The TableOfAuthorities object is a member of the TablesOfAuthorities collection. The TablesOfAuthorities collection includes all the tables of authorities in a document.

**Remarks:** Use TablesOfAuthorities (Index), where Index is the index number, to return a single TableOfAuthorities object. The index number represents the position of the table of authorities in the document. The following example includes category headers in the first table of authorities in the active document and then updates the table. Use the Add method to add a table of authorities to a document. The following example adds a table of authorities that includes all categories at the beginning of the active document.

## Properties (15)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TableOfAuthorities object.
- `Passim As Boolean  (read/write)`  
  True if five or more page references to the same authority are replaced with "Passim." Read/write Boolean.
- `KeepEntryFormatting As Boolean  (read/write)`  
  True if formatting from table of authorities entries is applied to the entries in the specified table of authorities. Read/write Boolean.
- `Category As Long  (read/write)`  
  Returns or sets the category of entries to be included in a table of authorities. Read/write Long.
- `Bookmark As String  (read/write)`  
  Returns or sets the name of the bookmark from which to collect table of authorities entries. Read/write String.
- `Separator As String  (read/write)`  
  Returns or sets up to five characters that appear between the sequence number and the page number in a table of authorities. Read/write String.
- `IncludeSequenceName As String  (read/write)`  
  Returns or sets the Sequence (SEQ) field identifier for a table of authorities. Read/write String.
- `EntrySeparator As String  (read/write)`  
  Returns or sets the characters (up to five) that separate a table of authorities entry and its page number. Read/write String.
- `PageRangeSeparator As String  (read/write)`  
  Returns or sets the characters (up to five) that separate a range of pages in a table of authorities. Read/write String.
- `IncludeCategoryHeader As Boolean  (read/write)`  
  True if the category name for a group of entries appears in the table of authorities. Read/write Boolean.
- `PageNumberSeparator As String  (read/write)`  
  Returns or sets the characters (up to five) that separate individual page references in a table of authorities. Read/write String.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within a table of authorities.
- `TabLeader As WdTabLeader  (read/write)`  
  Returns or sets the leader character that appears between entries and their associated page numbers in a table of authorities. Read/write WdTabLeader.

## Methods (2)

- `Delete()`  
  Deletes the specified table of authorities.
- `Update()`  
  Updates the entries shown in a table of authorities.
