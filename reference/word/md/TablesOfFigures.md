# TablesOfFigures

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020922-0000-0000-C000-000000000046}  

A collection of TableOfFigures objects that represent the tables of figures in a document.

**Remarks:** Use the TablesOfFigures property to return the TablesOfFigures collection. The following example applies the Classic format to all tables of figures in the active document. Use the Add method to add a table of figures to a document. A table of figures lists figure captions in the order in which they appear in the document. The following example replaces the selection in the active document with a table of figures that includes caption labels and page numbers. Use TablesOfFigures (Index), where Index is the index number, to return a single TableOfFigures object. The index number represents the position of the table of figures in the document. The following example updates the page numbers of the items in the first table of figures in the active document.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TablesOfFigures object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of tables of figures in the collection. Read-only.
- `Format As WdTofFormat  (read/write)`  
  Returns or sets the formatting for the tables of figures in the specified document. Read/write WdTofFormat.

## Methods (3)

- `Item(Index As Long) As TableOfFigures`  
  Returns an individual TableOfFigures object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `MarkEntry(Range As Range, [Entry As Variant], [EntryAutoText As Variant], [TableID As Variant], [Level As Variant]) As Field`  
  Inserts a TC (Table of Contents Entry) field after the specified range. The method returns a Field object representing the TC field.
    - `Range As Range` (required): The location of the entry. The TC field is inserted after Range.
    - `Entry As Variant` (optional): The text that appears in the table of contents or table of figures. To indicate a subentry, include the main entry text and the subentry text, separated by a colon (:) (for example, "Introduction:The Product").
    - `EntryAutoText As Variant` (optional): The AutoText entry name that includes text for the index, table of figures, or table of contents (Entry is ignored).
    - `TableID As Variant` (optional): A one-letter identifier for the table of figures or table of contents item (for example, "i" for an "illustration").
    - `Level As Variant` (optional): A level for the entry in the table of contents or table of figures.
- `Add(Range As Range, [Caption As Variant], [IncludeLabel As Variant], [UseHeadingStyles As Variant], [UpperHeadingLevel As Variant], [LowerHeadingLevel As Variant], [UseFields As Variant], [TableID As Variant], [RightAlignPageNumbers As Variant], [IncludePageNumbers As Variant], [AddedStyles As Variant], [UseHyperlinks As Variant], [HidePageNumbersInWeb As Variant]) As TableOfFigures`  
  Returns a TableOfFigures object that represents a table of figures added to a document.
    - `Range As Range` (required): The range where you want the table of figures to appear.
    - `Caption As Variant` (optional): The label that identifies the items you want to include in the table of figures. Corresponds to the \c switch for a Table of Contents (TOC) field. The default value is "Figure."
    - `IncludeLabel As Variant` (optional): True to include the caption label and caption number in the table of figures. The default value is True.
    - `UseHeadingStyles As Variant` (optional): True to use built-in heading styles to create the table of figures. The default value is False.
    - `UpperHeadingLevel As Variant` (optional): The starting heading level for the table of figures, if UseHeadingStyles is set to True. Corresponds to the starting value used with the \o switch for a Table of Contents (TOC) field. The default value is 1.
    - `LowerHeadingLevel As Variant` (optional): The ending heading level for the table of figures, if UseHeadingStyles is set to True. Corresponds to the ending value used with the \o switch for a Table of Contents (TOC) field. The default value is 9.
    - `UseFields As Variant` (optional): True to use Table of Contents Entry (TC) fields to create the table of figures. Use the MarkEntry method to mark entries you want to include in the table of figures. The default value is False.
    - `TableID As Variant` (optional): A one-letter identifier that's used to build a table of figures from Table of Contents Entry (TC) fields. Corresponds to the \f switch for a Table of Contents (TOC) field. For example, "i" builds a table of figures for an illustration.
    - `RightAlignPageNumbers As Variant` (optional): True align page numbers with the right margin in the table of figures. The default value is True.
    - `IncludePageNumbers As Variant` (optional): True if page numbers are included in the table of figures. The default value is True.
    - `AddedStyles As Variant` (optional): The string name for additional styles used to compile the table of figures (styles other than the Heading 1 &ndash; Heading 9 styles).
    - `UseHyperlinks As Variant` (optional): True if entries in a table of figures should be formatted as hyperlinks when publishing to the Web. The default value is True.
    - `HidePageNumbersInWeb As Variant` (optional): True if page numbers in a table of figures should be hidden when publishing to the Web. The default value is True.
