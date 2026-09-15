# Index

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002097D-0000-0000-C000-000000000046}  

Represents a single index. The Index object is a member of the Indexes collection. The Indexes collection includes all the indexes in the specified document.

**Remarks:** Use Indexes (Index), where Index is the index number, to return a single Index object. The index number represents the position of the Index object in the document. The following example updates the first index in the active document. Use the Add method to create an index and add it to the Indexes collection. The following example creates an index at the end of the active document.

## Properties (13)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Index object.
- `HeadingSeparator As WdHeadingSeparator  (read/write)`  
  Returns or sets the text between alphabetical groups (entries that start with the same letter) in the index. Corresponds to the \h switch for an INDEX field. Read/write WdHeadingSeparator.
- `RightAlignPageNumbers As Boolean  (read/write)`  
  True if page numbers are aligned with the right margin in an index. Read/write Boolean.
- `Type As WdIndexType  (read/write)`  
  Returns or sets the index type. Read/write WdIndexType.
- `NumberOfColumns As Long  (read/write)`  
  Sets or returns the number of columns for each page of an index. Read/write Long.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within an index.
- `TabLeader As WdTabLeader  (read/write)`  
  Returns or sets the leader character between entries in an index and their associated page numbers. Read/write WdTabLeader.
- `AccentedLetters As Boolean  (read/write)`  
  True if the specified index contains separate headings for accented letters (for example, words that begin with "" are under one heading and words that begin with "A" are under another). Read/write Boolean.
- `SortBy As WdIndexSortBy  (read/write)`  
  Returns or sets the sorting criteria for the specified index. Read/write WdIndexSortBy.
- `Filter As WdIndexFilter  (read/write)`  
  Returns or sets a value that specifies how Microsoft Word classifies the first character of entries in the specified index.read/write Long. Can be one of the following wdIndexFilter constants.
- `IndexLanguage As WdLanguageID  (read/write)`  
  Returns or sets a WdLanguageID constant that represents the sorting language to use for the specified index. Read/write .

## Methods (2)

- `Delete()`  
  Deletes the specified index.
- `Update()`  
  Updates the entries shown in specified index.
