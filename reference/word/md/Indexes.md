# Indexes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002097C-0000-0000-C000-000000000046}  

A collection of Index objects that represents all the indexes in the specified document.

**Remarks:** Use the Indexes property to return the Indexes collection. The following example formats indexes in the active document with the classic format. Use the Add method to create an index and add it to the Indexes collection. The following example creates an index at the end of the active document. Use Indexes (Index), where Index is the index number, to return a single Index object. The index number represents the position of the Index object in the document. The following example updates the first index in the active document.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Indexes object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of indexes in the collection. Read-only.
- `Format As WdIndexFormat  (read/write)`  
  Returns or sets a WdIndexFormat that represents the formatting for the indexes in the specified document. Read/write.

## Methods (5)

- `Item(Index As Long) As Index`  
  Returns an individual Index object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `MarkEntry(Range As Range, [Entry As Variant], [EntryAutoText As Variant], [CrossReference As Variant], [CrossReferenceAutoText As Variant], [BookmarkName As Variant], [Bold As Variant], [Italic As Variant], [Reading As Variant]) As Field`  
  Inserts an XE (Index Entry) field after the specified range. The method returns a Field object representing the XE field.
    - `Range As Range` (required): The location of the entry. The XE field is inserted after Range.
    - `Entry As Variant` (optional): The text that appears in the index. To indicate a subentry, include the main entry text and the subentry text, separated by a colon (:) (for example, "Introduction:The Product").
    - `EntryAutoText As Variant` (optional): The AutoText entry name that includes text for the index, table of figures, or table of contents (Entry is ignored).
    - `CrossReference As Variant` (optional): A cross-reference that will appear in the index (for example, "See Apples").
    - `CrossReferenceAutoText As Variant` (optional): The AutoText entry name that contains the text for a cross-reference (CrossReference is ignored).
    - `BookmarkName As Variant` (optional): The name of the bookmark that marks the range of pages you want to appear in the index. If this argument is omitted, the number of the page containing the XE field appears in the index.
    - `Bold As Variant` (optional): True to add bold formatting to the entry page numbers in the index.
    - `Italic As Variant` (optional): True to add italic formatting to the entry page numbers in the index.
    - `Reading As Variant` (optional): True shows an index entry in the right location when indexes are sorted phonetically (East Asian languages only).
- `MarkAllEntries(Range As Range, [Entry As Variant], [EntryAutoText As Variant], [CrossReference As Variant], [CrossReferenceAutoText As Variant], [BookmarkName As Variant], [Bold As Variant], [Italic As Variant])`  
  Inserts an XE (Index Entry) field after all instances of the text in Range.
    - `Range As Range` (required): The range whose text is marked with an XE field throughout the document.
    - `Entry As Variant` (optional): The text you want to appear in the index, in the form MainEntry[:Subentry].
    - `EntryAutoText As Variant` (optional): The AutoText entry that contains the text you want to appear in the index (if this argument is specified, Entry is ignored).
    - `CrossReference As Variant` (optional): A cross-reference that will appear in the index.
    - `CrossReferenceAutoText As Variant` (optional): The name of the AutoText entry that contains the text for a cross-reference (if this argument is specified, CrossReference is ignored).
    - `BookmarkName As Variant` (optional): The bookmark name that marks the range of pages you want to appear in the index. If this argument is omitted, the number of the page that contains the XE field appears in the index.
    - `Bold As Variant` (optional): True to add bold formatting to page numbers for index entries.
    - `Italic As Variant` (optional): True to add italic formatting to page numbers for index entries.
- `AutoMarkEntries(ConcordanceFileName As String)`  
  Automatically adds XE (Index Entry) fields to the specified document, using the entries from a concordance file.
    - `ConcordanceFileName As String` (required): The concordance file name that includes a list of items to be indexed.
- `Add(Range As Range, [HeadingSeparator As Variant], [RightAlignPageNumbers As Variant], [Type As Variant], [NumberOfColumns As Variant], [AccentedLetters As Variant], [SortBy As Variant], [IndexLanguage As Variant]) As Index`  
  Returns an Index object that represents a new index added to a document.
    - `Range As Range` (required): The range where you want the index to appear. The index replaces the range, if the range is not collapsed.
    - `HeadingSeparator As Variant` (optional): The text between alphabetical groups (entries that start with the same letter) in the index. Can be one of the WdHeadingSeparator constants.
    - `RightAlignPageNumbers As Variant` (optional): True to align page numbers with the right margin.
    - `Type As Variant` (optional): Specifies whether subentries are on the same line (run-in) as the main entry or on a separate line (indented) from the main entry. Can be either of the following WdIndexType constants: wdIndexIndent or wdIndexRunin.
    - `NumberOfColumns As Variant` (optional): The number of columns for each page of the index. Specifying 0 (zero) sets the number of columns in the index to the same number as in the document.
    - `AccentedLetters As Variant` (optional): True to include separate headings for accented letters in the index (for example, words that begin with "?" and words that begin with "A" are listed under separate headings).
    - `SortBy As Variant` (optional): The sorting criteria to be used for the specified index. Can be either of the following WdIndexSortBy constants: wdIndexSortByStroke or wdIndexSortBySyllable.
    - `IndexLanguage As Variant` (optional): The sorting language to be used for the specified index. Can be any of the WdLanguageID constants. For the list of valid WdLanguageID constants, see the Object Browser in the Visual Basic Editor.
