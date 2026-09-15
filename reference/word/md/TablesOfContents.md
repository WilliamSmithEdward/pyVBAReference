# TablesOfContents

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020914-0000-0000-C000-000000000046}  

A collection of TableOfContents objects that represent the tables of contents in a document.

**Remarks:** Use the TablesOfContents property to return the TablesOfContents collection. The following example inserts a table of contents entry that references the selected text in the active document. Use the Add method to add a table of contents to a document. The following example adds a table of contents at the beginning of the active document. The example builds the table of contents from all paragraphs styled as either Heading 1, Heading 2, or Heading 3. Use TablesOfContents (Index), where Index is the index number, to return a single TableOfContents object. The index number represents the position of the table of contents in the document. The following example updates the page numbers of the items in the first table of figures in the active document.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TablesOfContents object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of tables of contents in the collection. Read-only.
- `Format As WdTocFormat  (read/write)`  
  Returns or sets the formatting for the tables of contents in the specified document. Read/write WdTocFormat.

## Methods (3)

- `Item(Index As Long) As TableOfContents`  
  Returns an individual TableOfContents object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `MarkEntry(Range As Range, [Entry As Variant], [EntryAutoText As Variant], [TableID As Variant], [Level As Variant]) As Field`  
  Inserts a TC (Table of Contents Entry) field after the specified range. The method returns a Field object representing the TC field.
    - `Range As Range` (required): The location of the entry. The TC field is inserted after Range.
    - `Entry As Variant` (optional): The text that appears in the table of contents or table of figures. To indicate a subentry, include the main entry text and the subentry text, separated by a colon (:) (for example, "Introduction:The Product").
    - `EntryAutoText As Variant` (optional): The AutoText entry name that includes text for the index, table of figures, or table of contents (Entry is ignored).
    - `TableID As Variant` (optional): A one-letter identifier for the table of figures or table of contents item (for example, "i" for an "illustration").
    - `Level As Variant` (optional): A level for the entry in the table of contents or table of figures.
- `Add(Range As Range, [UseHeadingStyles As Variant], [UpperHeadingLevel As Variant], [LowerHeadingLevel As Variant], [UseFields As Variant], [TableID As Variant], [RightAlignPageNumbers As Variant], [IncludePageNumbers As Variant], [AddedStyles As Variant], [UseHyperlinks As Variant], [HidePageNumbersInWeb As Variant], [UseOutlineLevels As Variant]) As TableOfContents`  
  Returns a TableOfContents object that represents a table of contents added to a document.
    - `Range As Range` (required): The range where you want the table of contents to appear. The table of contents replaces the range, if the range isn't collapsed.
    - `UseHeadingStyles As Variant` (optional): True to use built-in heading styles to create the table of contents. The default value is True.
    - `UpperHeadingLevel As Variant` (optional): The starting heading level for the table of contents. Corresponds to the starting value used with the \o switch for a Table of Contents (TOC) field. The default value is 1.
    - `LowerHeadingLevel As Variant` (optional): The ending heading level for the table of contents. Corresponds to the ending value used with the \o switch for a Table of Contents (TOC) field. The default value is 9.
    - `UseFields As Variant` (optional): True if Table of Contents Entry (TC) fields are used to create the table of contents. Use the MarkEntry method to mark entries to be included in the table of contents. The default value is False.
    - `TableID As Variant` (optional): A one-letter identifier that's used to build a table of contents from TC fields. Corresponds to the \f switch for a Table of Contents (TOC) field. For example, "T" builds a table of contents from TC fields using the table identifier T. If this argument is omitted, TC fields aren't used.
    - `RightAlignPageNumbers As Variant` (optional): True if page numbers in the table of contents are aligned with the right margin. The default value is True.
    - `IncludePageNumbers As Variant` (optional): True to include page numbers in the table of contents. The default value is True.
    - `AddedStyles As Variant` (optional): The string name for additional styles used to compile the table of contents (styles other than the Heading 1 &ndash; Heading 9 styles). Use the Add method of a HeadingStyles object to create new heading styles.
    - `UseHyperlinks As Variant` (optional): True if entries in a table of contents should be formatted as hyperlinks when the document is being publishing to the web. The default value is True.
    - `HidePageNumbersInWeb As Variant` (optional): True if page numbers in a table of contents should be hidden when the document is being publishing to the web. The default value is True.
    - `UseOutlineLevels As Variant` (optional): True to use outline levels to create the table of contents. The default is False.
