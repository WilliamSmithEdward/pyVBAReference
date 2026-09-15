# TablesOfAuthorities

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020912-0000-0000-C000-000000000046}  

A collection of TableOfAuthorities objects (TOA fields) that represents the tables of authorities in a document.

**Remarks:** Use the TablesOfAuthorities property to return the TablesOfAuthorities collection. The following example applies the Classic built-in format to all the tables of authorities in the active document. Use the Add method to add a table of authorities to a document. A table of authorities is built from TA (Table of Authorities Entry) fields in a document. The following example adds a table of authorities that includes all categories at the beginning of the active document. Use TablesOfAuthorities (Index), where Index is the index number, to return a single TableOfAuthorities object. The index number represents the position of the table of authorities in the document. The following example includes category headers in the first table of authorities in the active document and then updates the table.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TablesOfAuthorities object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of tables of authorities in the collection. Read-only.
- `Format As WdToaFormat  (read/write)`  
  Returns or sets the formatting for the tables of authorities in the specified document. Read/write WdToaFormat.

## Methods (5)

- `Item(Index As Long) As TableOfAuthorities`  
  Returns an individual TableOfAuthorities object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add(Range As Range, [Category As Variant], [Bookmark As Variant], [Passim As Variant], [KeepEntryFormatting As Variant], [Separator As Variant], [IncludeSequenceName As Variant], [EntrySeparator As Variant], [PageRangeSeparator As Variant], [IncludeCategoryHeader As Variant], [PageNumberSeparator As Variant]) As TableOfAuthorities`  
  Returns a TableOfAuthorities object that represents a table of authorities added to a document.
    - `Range As Range` (required): The range where you want the table of authorities to appear. The table of authorities replaces the range, if the range isn't collapsed.
    - `Category As Variant` (optional): The category of entries you want to include in the table of authorities. Corresponds to the \c switch for a Table of Authorities (TOA) field. Values 0 through 16 correspond to the items listed in the Category box on the Table of Authorities tab in the Index and Tables dialog box (Reference command, Insert menu). The default value is 1.
    - `Bookmark As Variant` (optional): The string name of the bookmark from which you want to collect entries for the table of authorities. If Bookmark is specified, the entries are collected only from the portion of the document marked by the bookmark. Corresponds to the \b switch for a Table of Authorities (TOA) field.
    - `Passim As Variant` (optional): True to replace five or more page references to the same authority with Passim in the table of authorities. Corresponds to the \p switch for a Table of Authorities (TOA) field. If this argument is omitted, Passim is assumed to be False.
    - `KeepEntryFormatting As Variant` (optional): True to apply formatting from table of authorities entries to the entries in the table of authorities. Corresponds to the \f switch for a Table of Authorities (TOA) field. If this argument is omitted, KeepEntryFormatting is assumed to be True.
    - `Separator As Variant` (optional): The characters (up to five) between each sequence number and its page number in the table of authorities. Corresponds to the \d switch for a Table of Authorities (TOA) field. If argument is omitted, a hyphen (-) is used.
    - `IncludeSequenceName As Variant` (optional): A string that specifies the Sequence (SEQ) field identifier for the table of authorities. Corresponds to the \s switch for a Table of Authorities (TOA) field.
    - `EntrySeparator As Variant` (optional): The characters (up to five) that separate each entry and its page number in the table of authorities. Corresponds to the \e switch for a Table of Authorities (TOA) field. If this argument is omitted, no separator is used.
    - `PageRangeSeparator As Variant` (optional): The characters (up to five) that separate the beginning and ending page numbers in each page range the table of authorities. Corresponds to the \g switch for a Table of Authorities (TOA) field. If this argument is omitted, an en dash is used.
    - `IncludeCategoryHeader As Variant` (optional): True to have the category name for each group of entries appear in the table of authorities (for example, Cases). Corresponds to the \h switch for a Table of Authorities (TOA) field. If this argument is omitted, IncludeCategoryHeader is assumed to be True.
    - `PageNumberSeparator As Variant` (optional): The characters (up to five) that separate individual page numbers within page references in the table of authorities. Corresponds to the \l switch for a Table of Authorities (TOA) field. If this argument is omitted, a comma and a space are used.
- `NextCitation(ShortCitation As String)`  
  Finds and selects the next instance of the text specified by the ShortCitation parameter.
    - `ShortCitation As String` (required): The text of the short citation.
- `MarkCitation(Range As Range, ShortCitation As String, [LongCitation As Variant], [LongCitationAutoText As Variant], [Category As Variant]) As Field`  
  Inserts a TA (Table of Authorities Entry) field and returns the field as a Field object.
    - `Range As Range` (required): The location of the table of authorities entry. The TA field is inserted after Range.
    - `ShortCitation As String` (required): The short citation for the entry as it will appear in the Mark Citation dialog box (Insert menu, Index and Tables command).
    - `LongCitation As Variant` (optional): The long citation for the entry as it will appear in the table of authorities.
    - `LongCitationAutoText As Variant` (optional): The name of the AutoText entry that contains the text of the long citation as it will appear in the table of authorities.
    - `Category As Variant` (optional): The category number to be associated with the entry: 1 corresponds to the first category in the Category box in the Mark Citation dialog box, 2 corresponds to the second category, and so on.
- `MarkAllCitations(ShortCitation As String, [LongCitation As Variant], [LongCitationAutoText As Variant], [Category As Variant])`  
  Inserts a TA (Table of Authorities Entry) field after all instances of the ShortCitation text.
    - `ShortCitation As String` (required): The short citation for the entry as it will appear in the Mark Citation dialog box (Insert menu, Index and Tables command).
    - `LongCitation As Variant` (optional): The long citation string for the entry as it will appear in the table of authorities.
    - `LongCitationAutoText As Variant` (optional): The AutoText entry name that contains the text of the long citation as it will appear in the table of authorities.
    - `Category As Variant` (optional): The category number to be associated with the entry: 1 corresponds to the first category in the Category box in the Mark Citation dialog box, 2 corresponds to the second category, and so on.
