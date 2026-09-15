# Document

**Type:** Class  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020906-0000-0000-C000-000000000046}  

Represents a document. The Document object is a member of the Documents collection. The Documents collection contains all the Document objects that are currently open in Word.

**Remarks:** Use Documents (index), where index is the document name or index number, to return a single Document object. The following example closes the document named Report.doc without saving changes. The index number represents the position of the document in the Documents collection. The following example activates the first document in the Documents collection. Use the ActiveDocument property to refer to the document with the focus. The following example uses the Activate method to activate the document named Document 1. The example also sets the page orientation to landscape mode and then prints the document.

## Properties (193)

- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Document object.
- `BuiltInDocumentProperties As Object  (read-only)`  
  Returns a DocumentProperties collection that represents all the built-in document properties for the specified document.
- `CustomDocumentProperties As Object  (read-only)`  
  Returns a DocumentProperties collection that represents all the custom document properties for the specified document.
- `Path As String  (read-only)`  
  Returns the disk or Web path to the document. Read-only String.
- `Bookmarks As Bookmarks  (read-only)`  
  Returns a Bookmarks collection that represents all the bookmarks in a document. Read-only.
- `Tables As Tables  (read-only)`  
  Returns a Table collection that represents all the tables in the specified document. Read-only.
- `Footnotes As Footnotes  (read-only)`  
  Returns a Footnotes collection that represents all the footnotes in a document. Read-only.
- `Endnotes As Endnotes  (read-only)`  
  Returns an Endnotes collection that represents all the endnotes in a document. Read-only.
- `Comments As Comments  (read-only)`  
  Returns a Comments collection that represents all the comments in the specified document. Read-only.
- `Type As WdDocumentType  (read-only)`  
  Returns the document type (template or document). Read-only WdDocumentType.
- `AutoHyphenation As Boolean  (read/write)`  
  True if automatic hyphenation is turned on for the specified document. Read/write Boolean.
- `HyphenateCaps As Boolean  (read/write)`  
  True if words in all capital letters can be hyphenated. Read/write Boolean.
- `HyphenationZone As Long  (read/write)`  
  Returns or sets the width of the hyphenation zone, in points. Read/write Long.
- `ConsecutiveHyphensLimit As Long  (read/write)`  
  Returns or sets the maximum number of consecutive lines that can end with hyphens. Read/write. Long.
- `Sections As Sections  (read-only)`  
  Returns a Section collection that represents the sections in the specified document. Read-only.
- `Paragraphs As Paragraphs  (read-only)`  
  Returns a Paragraphs collection that represents all the paragraphs in the specified document. Read-only.
- `Words As Words  (read-only)`  
  Returns a Words collection that represents all the words in a document. Read-only.
- `Sentences As Sentences  (read-only)`  
  Returns a Sentences collection that represents all the sentences in the document. Read-only.
- `Characters As Characters  (read-only)`  
  Returns a Characters collection that represents the characters in a document. Read-only.
- `Fields As Fields  (read-only)`  
  Returns a Fields collection that represents all the fields in the document. Read-only.
- `FormFields As FormFields  (read-only)`  
  Returns a FormFields collection that represents all the form fields in the document. Read-only.
- `Styles As Styles  (read-only)`  
  Returns a Styles collection for the specified document. Read-only.
- `Frames As Frames  (read-only)`  
  Returns a Frames collection that represents all the frames in a document. Read-only.
- `TablesOfFigures As TablesOfFigures  (read-only)`  
  Returns a TablesOfFigures collection that represents the tables of figures in the specified document. Read-only.
- `Variables As Variables  (read-only)`  
  Returns a Variables collection that represents the variables stored in the specified document. Read-only.
- `MailMerge As MailMerge  (read-only)`  
  Returns a MailMerge object that represents the mail merge functionality for the specified document. Read-only.
- `Envelope As Envelope  (read-only)`  
  Returns an Envelope object that represents an envelope and envelope features in a document. Read-only.
- `FullName As String  (read-only)`  
  Returns a String that represents the name of a document, including the path. Read-only.
- `Revisions As Revisions  (read-only)`  
  Returns a Revisions collection that represents the tracked changes in the document or range. Read-only.
- `TablesOfContents As TablesOfContents  (read-only)`  
  Returns a TablesOfContents collection that represents the tables of contents in the specified document. Read-only.
- `TablesOfAuthorities As TablesOfAuthorities  (read-only)`  
  Returns a TableOfAuthorities collection that represents the tables of authorities in the specified document. Read-only.
- `PageSetup As PageSetup  (read/write)`  
  Returns a PageSetup object that is associated with the specified document.
- `Windows As Windows  (read-only)`  
  Returns a Windows collection that represents all windows for the specified document. Read-only.
- `TablesOfAuthoritiesCategories As TablesOfAuthoritiesCategories  (read-only)`  
  Returns a TablesOfAuthoritiesCategories collection that represents the available table of authorities categories for the specified document. Read-only.
- `Indexes As Indexes  (read-only)`  
  Returns an Indexes collection that represents all the indexes in the specified document. Read-only.
- `Saved As Boolean  (read/write)`  
  True if the specified document or template has not changed since it was last saved. False if Microsoft Word displays a prompt to save changes when the document is closed. Read/write Boolean.
- `Content As Range  (read-only)`  
  Returns a Range object that represents the main document story. Read-only.
- `ActiveWindow As Window  (read-only)`  
  Returns a Window object that represents the active window (the window with the focus). Read-only.
- `Kind As WdDocumentKind  (read/write)`  
  Returns or sets the format type that Microsoft Word uses when automatically formatting the specified document. Read/write WdDocumentKind.
- `ReadOnly As Boolean  (read-only)`  
  True if changes to the document cannot be saved to the original document. Read-only Boolean.
- `Subdocuments As Subdocuments  (read-only)`  
  Returns a Subdocuments collection that represents all the subdocuments in the specified document. Read-only.
- `IsMasterDocument As Boolean  (read-only)`  
  True if the specified document is a master document. Read-only Boolean.
- `DefaultTabStop As Single  (read/write)`  
  Returns or sets the interval (in points) between the default tab stops in the specified document. Read/write Single.
- `EmbedTrueTypeFonts As Boolean  (read/write)`  
  True if Microsoft Word embeds TrueType fonts in a document when it is saved. Read/write Boolean.
- `SaveFormsData As Boolean  (read/write)`  
  True if Microsoft Word saves the data entered in a form as a tab-delimited record for use in a database. Read/write Boolean.
- `ReadOnlyRecommended As Boolean  (read/write)`  
  True if Microsoft Word displays a message box whenever a user opens the document, suggesting that it be opened as read-only. Read/write Boolean.
- `SaveSubsetFonts As Boolean  (read/write)`  
  True if Microsoft Word saves a subset of the embedded TrueType fonts with the document. Read/write Boolean.
- `Compatibility As Boolean  (read/write)`  
  True if the compatibility option specified by the Type argument is enabled. Compatibility options affect how a document is displayed in Microsoft Word. Read/write Boolean.
- `StoryRanges As StoryRanges  (read-only)`  
  Returns a StoryRanges collection that represents all the stories in the specified document. Read-only.
- `CommandBars As CommandBars  (read-only)`  
  Returns a CommandBars collection that represents the menu bar and all the toolbars in Microsoft Word.
- `IsSubdocument As Boolean  (read-only)`  
  True if the specified document is a subdocument of a master document. Read-only Boolean.
- `SaveFormat As Long  (read-only)`  
  Returns the file format of the specified document or file converter. Read-only Long.
- `ProtectionType As WdProtectionType  (read-only)`  
  Returns the protection type for the specified document. Can be one of the following WdProtectionType constants: wdAllowOnlyComments, wdAllowOnlyFormFields, wdAllowOnlyReading, wdAllowOnlyRevisions, or wdNoProtection.
- `Hyperlinks As Hyperlinks  (read-only)`  
  Returns a Hyperlinks collection that represents all the hyperlinks in the specified document. Read-only.
- `Shapes As Shapes  (read-only)`  
  Returns a Shapes collection that represents all the Shape objects in the specified document. Read-only.
- `ListTemplates As ListTemplates  (read-only)`  
  Returns a ListTemplates collection that represents all the list formats for the specified document. Read-only.
- `Lists As Lists  (read-only)`  
  Returns a Lists collection that contains all the formatted lists in the specified document. Read-only.
- `UpdateStylesOnOpen As Boolean  (read/write)`  
  True if the styles in the specified document are updated to match the styles in the attached template each time the document is opened. Read/write Boolean.
- `AttachedTemplate As Variant  (read/write)`  
  Returns a Template object that represents the template attached to the specified document. Read/write Variant.
- `InlineShapes As InlineShapes  (read-only)`  
  Returns an InlineShapes collection that represents all the InlineShape objects in a document. Read-only.
- `Background As Shape  (read/write)`  
  Returns a Shape object that represents the background image for the specified document. Read-only.
- `GrammarChecked As Boolean  (read/write)`  
  True if a grammar check has been run on the specified range or document. Read/write Boolean.
- `SpellingChecked As Boolean  (read/write)`  
  True if spelling has been checked throughout the specified range or document. False if all or some of the range or document has not been checked for spelling. Read/write Boolean.
- `ShowGrammaticalErrors As Boolean  (read/write)`  
  True if grammatical errors are marked by a wavy green line in the specified document. Read/write Boolean.
- `ShowSpellingErrors As Boolean  (read/write)`  
  True if Microsoft Word underlines spelling errors in the document. Read/write Boolean.
- `PrintPostScriptOverText As Boolean  (read/write)`  
  True if PRINT field instructions (such as PostScript commands) in a document are to be printed on top of text and graphics when a PostScript printer is used. Read/write Boolean.
- `Container As Object  (read-only)`  
  Returns the object that represents the container application for the specified document. Read-only Object.
- `PrintFormsData As Boolean  (read/write)`  
  True if Microsoft Word prints onto a preprinted form only the data entered in the corresponding online form. Read/write Boolean.
- `ListParagraphs As ListParagraphs  (read-only)`  
  Returns a ListParagraphs object that represents all the numbered paragraphs in a document. Read-only.
- `Password As String  (write-only)`  
  Sets a password that must be supplied to open the specified document. Write-only String.
- `WritePassword As String  (write-only)`  
  Sets a password for saving changes to the specified document. Write-only String.
- `HasPassword As Boolean  (read-only)`  
  True if a password is required to open the specified document. Read-only Boolean.
- `WriteReserved As Boolean  (read-only)`  
  True if the specified document is protected with a write password. Read-only Boolean.
- `ActiveWritingStyle As String  (read/write)`  
  Returns or sets the writing style for a specified language in the specified document. Read/write String.
- `UserControl As Boolean  (read/write)`  
  True if the document was created or opened by the user. Read/write Boolean.
- `ReadabilityStatistics As ReadabilityStatistics  (read-only)`  
  Returns a ReadabilityStatistics collection that represents the readability statistics for the specified document or range. Read-only.
- `GrammaticalErrors As ProofreadingErrors  (read-only)`  
  Returns a ProofreadingErrors collection that represents the sentences that failed the grammar check in the specified document. Read-only.
- `SpellingErrors As ProofreadingErrors  (read-only)`  
  Returns a ProofreadingErrors collection that represents the words identified as spelling errors in the specified document or range. Read-only.
- `VBProject As VBProject  (read-only)`  
  Returns the VBProject object for the specified template or document.
- `FormsDesign As Boolean  (read-only)`  
  True if the specified document is in form design mode. Read-only Boolean.
- `_CodeName As String  (read/write)`
- `CodeName As String  (read-only)`  
  Returns the code name for the specified document. Read-only String.
- `SnapToGrid As Boolean  (read/write)`  
  True if AutoShapes or East Asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized in the specified document. Read/write Boolean.
- `SnapToShapes As Boolean  (read/write)`  
  True if Microsoft Word automatically aligns AutoShapes or East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other AutoShapes or East Asian characters in the specified document. Read/write Boolean.
- `GridDistanceHorizontal As Single  (read/write)`  
  Returns or sets a Single that represents the amount of horizontal space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize AutoShapes or East Asian characters in the specified document. Read/write.
- `GridDistanceVertical As Single  (read/write)`  
  Returns or sets a Single that represents the amount of vertical space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize AutoShapes or East Asian characters in the specified document. Read/write.
- `GridOriginHorizontal As Single  (read/write)`  
  Returns or sets a Single that represents the point, relative to the left edge of the page, where you want the invisible grid for drawing, moving, and resizing AutoShapes or East Asian characters to begin in the specified document. Read/write.
- `GridOriginVertical As Single  (read/write)`  
  Returns or sets a Single that represents the point, relative to the top of the page, where you want the invisible grid for drawing, moving, and resizing AutoShapes or East Asian characters to begin in the specified document. Read/write.
- `GridSpaceBetweenHorizontalLines As Long  (read/write)`  
  Returns or sets the interval at which Microsoft Word displays horizontal character gridlines in print layout view. Read/write Long.
- `GridSpaceBetweenVerticalLines As Long  (read/write)`  
  Returns or sets the interval at which Microsoft Word displays vertical character gridlines in print layout view. Read/write Long.
- `GridOriginFromMargin As Boolean  (read/write)`  
  True if Microsoft Word starts the character grid from the upper-left corner of the page. Read/write Boolean.
- `KerningByAlgorithm As Boolean  (read/write)`  
  True if Microsoft Word kerns half-width Latin characters and punctuation marks in the specified document. Read/write Boolean.
- `JustificationMode As WdJustificationMode  (read/write)`  
  Returns or sets the character spacing adjustment for the specified document. Read/write WdJustificationMode.
- `FarEastLineBreakLevel As WdFarEastLineBreakLevel  (read/write)`  
  Returns or sets a WdFarEastLineBreakLevel that represents the line break control level for the specified document. Read/write.
- `NoLineBreakBefore As String  (read/write)`  
  Returns or sets the kinsoku characters before which Microsoft Word will not break a line. Read/write String.
- `NoLineBreakAfter As String  (read/write)`  
  Returns or sets the kinsoku characters after which Microsoft Word will not break a line. Read/write String.
- `TrackRevisions As Boolean  (read/write)`  
  True if changes are tracked in the specified document. Read/write Boolean.
- `PrintRevisions As Boolean  (read/write)`  
  True if revision marks are printed with the document. False if revision marks aren't printed (that is, tracked changes are printed as if they'd been accepted). Read/write Boolean.
- `ActiveTheme As String  (read-only)`  
  Returns the name of the active theme plus the theme formatting options for the specified document. Read-only String.
- `ActiveThemeDisplayName As String  (read-only)`  
  Returns the display name of the active theme for the specified document. Read-only String.
- `Email As Email  (read-only)`  
  Returns an Email object that contains all the email-related properties of the current document. Read-only.
- `Scripts As Scripts  (read-only)`  
  Returns a Scripts collection that represents the collection of HTML scripts in the specified object.
- `LanguageDetected As Boolean  (read/write)`  
  Returns or sets a value that specifies whether Microsoft Word has detected the language of the specified text. Read/write Boolean.
- `FarEastLineBreakLanguage As WdFarEastLineBreakLanguageID  (read/write)`  
  Returns or sets a WdFarEastLineBreakLanguageID that represents the East Asian language to use when breaking lines of text in the specified document or template. Read/write.
- `Frameset As Frameset  (read-only)`  
  Returns a Frameset object that represents an entire frames page or a single frame on a frames page. Read-only.
- `ClickAndTypeParagraphStyle As Variant  (read/write)`  
  Returns or sets the default paragraph style applied to text by the Click and Type feature in the specified document. Read/write Variant.
- `WebOptions As WebOptions  (read-only)`  
  Returns the WebOptions object, which contains document-level attributes used by Microsoft Word when you save a document as a webpage or open a webpage. Read-only.
- `OpenEncoding As MsoEncoding  (read-only)`  
  Returns the encoding used to open the specified document. Read-only MsoEncoding.
- `SaveEncoding As MsoEncoding  (read/write)`  
  Returns or sets the encoding to use when saving a document. Read/write MsoEncoding.
- `OptimizeForWord97 As Boolean  (read/write)`  
  True if Microsoft Word optimizes the current document for viewing in Microsoft Word 97 by disabling any incompatible formatting. Read/write Boolean.
- `VBASigned As Boolean  (read-only)`  
  True if the Microsoft Visual Basic for Applications (VBA) project for the specified document has been digitally signed. Read-only Boolean.
- `MailEnvelope As MsoEnvelope  (read-only)`  
  Returns an MsoEnvelope object that represents an email header for a document.
- `DisableFeatures As Boolean  (read/write)`  
  True disables all features introduced after the version specified in the DisableFeaturesIntroducedAfter property. The default value is False. Read/write Boolean.
- `DoNotEmbedSystemFonts As Boolean  (read/write)`  
  True for Microsoft Word to not embed common system fonts. Read/write Boolean.
- `Signatures As SignatureSet  (read-only)`  
  Returns a SignatureSet collection that represents the digital signatures for a document.
- `DefaultTargetFrame As String  (read/write)`  
  Returns or sets a String indicating the browser frame in which to display a webpage reached through a hyperlink. Read/write.
- `HTMLDivisions As HTMLDivisions  (read-only)`  
  Returns an HTMLDivisions collection that represents the HTML DIV elements in a web document.
- `DisableFeaturesIntroducedAfter As WdDisableFeaturesIntroducedAfter  (read/write)`  
  Disables all features introduced after a specified version of Microsoft Word in the document only. Read/write WdDisableFeaturesIntroducedAfter.
- `RemovePersonalInformation As Boolean  (read/write)`  
  True if Microsoft Word removes all user information from comments, revisions, and the Properties dialog box upon saving a document. Read/write Boolean.
- `TextEncoding As MsoEncoding  (read/write)`  
  Returns or sets the code page, or character set, that Microsoft Word uses for a document saved as an encoded text file. Read/write MsoEncoding.
- `TextLineEnding As WdLineEndingType  (read/write)`  
  Returns or sets a WdLineEndingType constant indicating how Microsoft Word marks the line and paragraph breaks in documents saved as text files. Read/write.
- `StyleSheets As StyleSheets  (read-only)`  
  Returns a StyleSheets collection that represents the Web style sheets attached to a document.
- `DefaultTableStyle As Variant  (read-only)`  
  Returns a Variant that represents the table style that is applied to all newly created tables in a document. Read-only.
- `PasswordEncryptionProvider As String  (read-only)`  
  Returns a String specifying the name of the algorithm encryption provider that Microsoft Word uses when encrypting documents with passwords. Read-only.
- `PasswordEncryptionAlgorithm As String  (read-only)`  
  Returns a String indicating the algorithm Microsoft Word uses for encrypting documents with passwords. Read-only.
- `PasswordEncryptionKeyLength As Long  (read-only)`  
  Returns a Long indicating the key length of the algorithm Microsoft Word uses when encrypting documents with passwords. Read-only.
- `PasswordEncryptionFileProperties As Boolean  (read-only)`  
  True if Microsoft Word encrypts file properties for password-protected documents. Read-only Boolean.
- `EmbedLinguisticData As Boolean  (read/write)`  
  True for Microsoft Word to embed speech and handwriting so that data can be converted back to speech or handwriting. Read/write Boolean.
- `FormattingShowFont As Boolean  (read/write)`  
  True for Microsoft Word to display font formatting in the Styles and Formatting task pane. Read/write Boolean.
- `FormattingShowClear As Boolean  (read/write)`  
  True for Microsoft Word to show clear formatting in the Styles and Formatting task pane. Read/write Boolean.
- `FormattingShowParagraph As Boolean  (read/write)`  
  True for Microsoft Word to display paragraph formatting in the Styles and Formatting task pane. Read/write Boolean.
- `FormattingShowNumbering As Boolean  (read/write)`  
  True for Microsoft Word to display number formatting in the Styles and Formatting task pane. Read/write Boolean.
- `FormattingShowFilter As WdShowFilter  (read/write)`  
  Sets or returns a WdShowFilter constant that represents the styles and formatting displayed in the Styles and Formatting task pane. Read/write Boolean.
- `Permission As Permission  (read-only)`  
  Returns a Permission object that represents the permission settings in the specified document.
- `XMLSchemaReferences As XMLSchemaReferences  (read-only)`  
  Returns an XMLSchemaReferences collection that represents the schemas attached to a document.
- `SmartDocument As SmartDocument  (read-only)`  
  Returns a SmartDocument object that represents the settings for a smart document solution.
- `Sync As Sync  (read-only)`
- `EnforceStyle As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether formatting restrictions are enforced in a protected document.
- `AutoFormatOverride As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether automatic formatting options override formatting restrictions in a document where formatting restrictions are in effect.
- `XMLShowAdvancedErrors As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether error message text is generated from the built-in Microsoft Word error messages or from the Microsoft XML Core Services (MSXML) 5.0 component included with Office.
- `XMLUseXSLTWhenSaving As Boolean  (read/write)`  
  Returns a Boolean that represents whether to save a document through an Extensible Stylesheet Language Transformation (XSLT). True saves a document through an XSLT.
- `XMLSaveThroughXSLT As String  (read/write)`  
  Sets or returns a String that specifies the path and file name for the Extensible Stylesheet Language Transformation (XSLT) to apply when a user saves a document.
- `DocumentLibraryVersions As DocumentLibraryVersions  (read-only)`  
  Returns a DocumentLibraryVersions collection that represents the collection of versions of a shared document that has versioning enabled and that is stored in a document library on a server.
- `ReadingModeLayoutFrozen As Boolean  (read/write)`  
  Sets or returns a Boolean that represents whether pages displayed in reading layout view are frozen to a specified size for inserting handwritten markup into a document.
- `RemoveDateAndTime As Boolean  (read/write)`  
  Sets or returns a Boolean indicating whether a document stores the date and time metadata for tracked changes. .
- `ReadingLayoutSizeX As Long  (read/write)`  
  Sets or returns a Long that represents the width of pages in a document when it is displayed in reading layout view and is frozen for entering handwritten markup.
- `ReadingLayoutSizeY As Long  (read/write)`  
  Sets or returns a Long that represents the height of pages in a document when it is displayed in reading layout view and is frozen for entering handwritten markup.
- `StyleSortMethod As WdStyleSort  (read/write)`  
  Returns or sets a WdStyleSort constant that represents the sort method to use when sorting styles in the Styles task pane. Read/write.
- `ContentTypeProperties As MetaProperties  (read-only)`  
  Returns a MetaProperties collection that represents the metadata stored in a document, such as author name, subject, and company. Read-only.
- `TrackMoves As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to mark moved text when Track Changes is turned on. Read/write.
- `TrackFormatting As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to track formatting changes when change tracking is turned on. Read/write.
- `OMaths As OMaths  (read-only)`  
  Returns an OMaths collection that represents the OMath objects within the specified range. Read-only.
- `ServerPolicy As ServerPolicy  (read-only)`  
  Returns a ServerPolicy object that represents a policy specified for a document stored on a server running Microsoft Office SharePoint Server 2007. Read-only.
- `ContentControls As ContentControls  (read-only)`  
  Returns a ContentControls collection that represents all the content controls in a document. Read-only.
- `DocumentInspectors As DocumentInspectors  (read-only)`  
  Returns a DocumentInspectors collection that enables you to locate hidden personal information, such as author name, company name, and revision date. Read-only.
- `Bibliography As Bibliography  (read-only)`  
  Returns a Bibliography object that represents the bibliography references contained within a document. Read-only.
- `LockTheme As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether a user can change a document theme. Read/write.
- `LockQuickStyleSet As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether users can change which set of Quick Styles is being used. Read/write.
- `OriginalDocumentTitle As String  (read-only)`  
  Returns a String that represents the document title for the original document after running a legal-blackline document compare function. Read-only.
- `RevisedDocumentTitle As String  (read-only)`  
  Returns a String that represents the document title for a revised document after running a legal-blackline document compare function. Read-only.
- `CustomXMLParts As CustomXMLParts  (read-only)`  
  Returns a CustomXMLParts collection that represents the custom XML in the XML data store. Read-only.
- `FormattingShowNextLevel As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Microsoft Word shows the next heading level when the previous heading level is used. Read/write.
- `FormattingShowUserStyleName As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to show user-defined styles. Read/write.
- `Research As Research  (read-only)`  
  Returns a Research object that represents the research service for a document. Read-only.
- `Final As Boolean  (read/write)`  
  Returns or sets a Boolean that indicates whether a document is final. Read/write.
- `OMathBreakBin As WdOMathBreakBin  (read/write)`  
  Returns or sets a WdOMathBreakBin constant that represents where Microsoft Word places binary operators when equations span two or more lines. Read/write.
- `OMathBreakSub As WdOMathBreakSub  (read/write)`  
  Returns or sets a WdOMathBreakSub constant that represents how Microsoft Word handles a subtraction operator that falls before a line break. Read/write.
- `OMathJc As WdOMathJc  (read/write)`  
  Returns or sets a WdOMathJc constant that represents the default justification-left, right, centered, or centered as a group-of a group of equations. Read/write.
- `OMathLeftMargin As Single  (read/write)`  
  Returns or sets a Single that represents the left margin for equations. Read/write.
- `OMathRightMargin As Single  (read/write)`  
  Returns or sets a Single that represents the right margin for equations. Read/write.
- `OMathWrap As Single  (read/write)`  
  Returns or sets a Single that represents the placement of the second line of an equation that wraps to a new line. Read/write.
- `OMathIntSubSupLim As Boolean  (read/write)`  
  Returns or sets a Boolean that represents the default location of limits for integrals. Read/write.
- `OMathNarySupSubLim As Boolean  (read/write)`  
  Returns or sets a Boolean that represents the default location of limits for n-ary objects other than integrals. Read/write.
- `OMathSmallFrac As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to use small fractions in equations contained within the document. Read/write.
- `WordOpenXML As String  (read-only)`  
  Returns a String that represents the flat XML format for the Word Open XML contents of the document. Read-only.
- `DocumentTheme As OfficeTheme  (read-only)`  
  Returns an OfficeTheme object that represents the Microsoft Office theme applied to a document. Read-only.
- `HasVBProject As Boolean  (read-only)`  
  Returns a Boolean that represents whether a document has an attached Microsoft Visual Basic for Applications project. Read-only.
- `OMathFontName As String  (read/write)`  
  Returns or sets a String that represents the name of the font used in a document to display equations. Read/write.
- `EncryptionProvider As String  (read/write)`  
  Returns a String specifying the name of the algorithm encryption provider that Microsoft Word uses when encrypting documents. Read/write.
- `UseMathDefaults As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to use the default math settings when creating new equations. Read/write.
- `CurrentRsid As Long  (read-only)`  
  Returns a Long that represents a random number that Word assigns to changes in a document. Read-only.
- `CompatibilityMode As Long  (read-only)`  
  Returns a Long that specifies the compatibility mode that Word uses when opening the document. Read-only.
- `CoAuthoring As CoAuthoring  (read-only)`  
  Returns a CoAuthoring object that provides the entry point into the co authoring object model. Read-only.
- `Broadcast As Broadcast  (read-only)`  
  Returns a Broadcast object that represents a broadcast session, in which presenters can present Word documents to remote participants over the web without the participants needing to have rich clients installed.
- `ChartDataPointTrack As Boolean  (read/write)`  
  Returns or sets a Boolean that specifies whether charts in the active document use cell-reference data-point tracking. Read/write.
- `IsInAutosave As Boolean  (read-only)`  
  Returns False if the most recent firing of the Application.DocumentBeforeSave event (Word) event was the result of a manual save by the user, and not an automatic save. Read-only.
- `WorkIdentity As String  (read/write)`
- `AutoSaveOn As Boolean  (read/write)`  
  True if the edits in the document are automatically saved. Read/write Boolean.
- `SensitivityLabel As ISensitivityLabel  (read-only)`  
  Returns the Microsoft Office SensitivityLabel object from the Document.
- `TrackJustMyRevisions As Boolean  (read/write)`

## Methods (97)

- `Close([SaveChanges As Variant], [OriginalFormat As Variant], [RouteDocument As Variant])`  
  Closes the specified document.
    - `SaveChanges As Variant` (optional): Specifies the save action for the document. Can be one of the following WdSaveOptions constants: wdDoNotSaveChanges, wdPromptToSaveChanges, or wdSaveChanges.
    - `OriginalFormat As Variant` (optional): Specifies the save format for the document. Can be one of the following WdOriginalFormat constants: wdOriginalDocumentFormat, wdPromptUser, or wdWordDocument.
    - `RouteDocument As Variant` (optional): True to route the document to the next recipient. If the document does not have a routing slip attached, this argument is ignored.
- `Repaginate()`  
  Repaginates the entire document.
- `FitToPages()`  
  Decreases the font size of text just enough so that the document will fit on one fewer pages.
- `ManualHyphenation()`  
  Initiates manual hyphenation of a document, one line at a time.
- `Select()`  
  Selects the contents of the specified document.
- `DataForm()`  
  Displays the Data Form dialog box, in which you can add, delete, or modify records.
- `Save()`  
  Saves the specified document.
- `SendMail()`  
  Opens a message window for sending the specified document through Microsoft Exchange.
- `Range([Start As Variant], [End As Variant]) As Range`  
  Returns a Range object by using the specified starting and ending character positions.
    - `Start As Variant` (optional): The starting character position.
    - `End As Variant` (optional): The ending character position.
- `RunAutoMacro(Which As WdAutoMacros)`  
  Runs an auto macro that's stored in the specified document. If the specified auto macro doesn't exist, nothing happens.
    - `Which As WdAutoMacros` (required): The auto macro to run.
- `Activate()`  
  Activates the specified document so that it becomes the active document.
- `PrintPreview()`  
  Switches the view to print preview.
- `GoTo([What As Variant], [Which As Variant], [Count As Variant], [Name As Variant]) As Range`  
  Returns a Range object that represents the start position of the specified item, such as a page, bookmark, or field.
    - `What As Variant` (optional): The kind of item to which the range or selection is moved. Can be one of the WdGoToItem constants.
    - `Which As Variant` (optional): The item to which the range or selection is moved. Can be one of the WdGoToDirection constants.
    - `Count As Variant` (optional): The number of the item in the document. The default value is 1. Only positive values are valid. To specify an item that precedes the range or selection, use wdGoToPrevious as the Which argument and specify a value for the Count value.
    - `Name As Variant` (optional): If the What argument is wdGoToBookmark, wdGoToComment, wdGoToField, or wdGoToObject, this argument specifies a name. Only positive values are valid. To specify an item that precedes the range or selection, use wdGoToPrevious as the Which argument and specify a value for the Count argument.
- `Undo([Times As Variant]) As Boolean`  
  Undoes the last action or a sequence of actions, which are displayed in the Undo list. Returns True if the actions were successfully undone.
    - `Times As Variant` (optional): The number of actions to be undone.
- `Redo([Times As Variant]) As Boolean`  
  Redoes the last action that was undone (reverses the Undo method). Returns True if the actions were redone successfully.
    - `Times As Variant` (optional): The number of actions to be redone.
- `ComputeStatistics(Statistic As WdStatistic, [IncludeFootnotesAndEndnotes As Variant]) As Long`  
  Returns a statistic based on the contents of the specified document. Long.
    - `Statistic As WdStatistic` (required): The statistic to compute.
    - `IncludeFootnotesAndEndnotes As Variant` (optional): True to include footnotes and endnotes when computing statistics. If this argument is omitted, the default value is False.
- `MakeCompatibilityDefault()`  
  Sets the compatibility options.
- `Unprotect([Password As Variant])`  
  Removes protection from the specified document. .
    - `Password As Variant` (optional): The password string used to protect the document. Passwords are case-sensitive. If the document is protected with a password and the correct password isn't supplied, a dialog box prompts the user for the password.
- `RunLetterWizard([LetterContent As Variant], [WizardMode As Variant])`  
  Runs the Letter Wizard on the specified document.
    - `LetterContent As Variant` (optional): A LetterContent object. Any filled properties in the LetterContent object show up as prefilled elements in the Letter Wizard dialog boxes. If this argument is omitted, the GetLetterContent method is automatically used to get a LetterContent object from the specified document.
    - `WizardMode As Variant` (optional): True to display the Letter Wizard dialog box as a series of steps with a Next, Back, and Finish button. False to display the Letter Wizard dialog box as if it were opened from the Tools menu (a properties dialog box with an OK button and a Cancel button). The default value is True.
- `GetLetterContent() As LetterContent`  
  Retrieves letter elements from the specified document and returns a LetterContent object.
- `SetLetterContent(LetterContent As Variant)`  
  Inserts the contents of the specified LetterContent object into a document.
    - `LetterContent As Variant` (required): The that includes the various elements of the letter.
- `CopyStylesFromTemplate(Template As String)`  
  Copies styles from the specified template to a document.
    - `Template As String` (required): The template file name.
- `UpdateStyles()`  
  Copies all styles from the attached template into the document, overwriting any existing styles in the document that have the same name.
- `CheckGrammar()`  
  Begins a spelling and grammar check for the specified document or range.
- `CheckSpelling([CustomDictionary As Variant], [IgnoreUppercase As Variant], [AlwaysSuggest As Variant], [CustomDictionary2 As Variant], [CustomDictionary3 As Variant], [CustomDictionary4 As Variant], [CustomDictionary5 As Variant], [CustomDictionary6 As Variant], [CustomDictionary7 As Variant], [CustomDictionary8 As Variant], [CustomDictionary9 As Variant], [CustomDictionary10 As Variant])`  
  Begins a spelling check for the specified document or range. .
    - `IgnoreUppercase As Variant` (optional): True if capitalization is ignored. If this argument is omitted, the current value of the IgnoreUppercase property is used.
    - `AlwaysSuggest As Variant` (optional): True for Microsoft Word to always suggest alternative spellings. If this argument is omitted, the current value of the SuggestSpellingCorrections property is used.
- `FollowHyperlink([Address As Variant], [SubAddress As Variant], [NewWindow As Variant], [AddHistory As Variant], [ExtraInfo As Variant], [Method As Variant], [HeaderInfo As Variant])`  
  Displays a cached document, if it has already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.
    - `Address As Variant` (optional): The address of the target document.
    - `SubAddress As Variant` (optional): The location within the target document. The default value is an empty string.
    - `NewWindow As Variant` (optional): True to display the target location in a new window. The default value is False.
    - `AddHistory As Variant` (optional): True to add the link to the current day's history folder.
    - `ExtraInfo As Variant` (optional): A string or a byte array that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use ExtraInfo to specify the coordinates of an image map, the contents of a form, or a FAT file name. The string is either posted or appended, depending on the value of Method. Use the ExtraInfoRequired property to determine whether extra information is required.
    - `Method As Variant` (optional): Specifies the way additional information for HTTP is handled. Can be any MsoExtraInfoMethod constant.
    - `HeaderInfo As Variant` (optional): A string that specifies header information for the HTTP request. The default value is an empty string.You can combine several header lines into a single string by using the following syntax: "string1" & vbCr & "string2". The specified string is automatically converted into ANSI characters. Note that the HeaderInfo argument may overwrite default HTTP header fields.
- `AddToFavorites()`  
  Creates a shortcut to the document or hyperlink and adds it to the Favorites folder.
- `Reload()`  
  Reloads a cached document by resolving the hyperlink to the document and downloading it.
- `RemoveNumbers([NumberType As Variant])`  
  Removes numbers or bullets from the specified document.
    - `NumberType As Variant` (optional): The type of number to be removed.
- `ConvertNumbersToText([NumberType As Variant])`  
  Changes the list numbers and LISTNUM fields in the specified Document to text. .
- `CountNumberedItems([NumberType As Variant], [Level As Variant]) As Long`  
  Returns the number of bulleted or numbered items and LISTNUM fields in the specified Document object.
    - `NumberType As Variant` (optional): The type of numbers to be counted. Can be one of the WdNumberType constants. The default value is wdNumberAllNumbers.
    - `Level As Variant` (optional): A number that corresponds to the numbering level you want to count. If this argument is omitted, all levels are counted.
- `Post()`  
  Posts the specified document to a public folder in Microsoft Exchange. .
- `ToggleFormsDesign()`  
  Switches form design mode on or off.
- `GetCrossReferenceItems(ReferenceType As Variant) As Variant`  
  Returns an array of items that can be cross-referenced based on the specified cross-reference type.
    - `ReferenceType As Variant` (required): The type of item you want to insert a cross-reference to. Can be any WdReferenceType constant.
- `AutoFormat()`  
  Automatically formats a document.
- `ViewCode()`  
  Displays the code window for the selected Microsoft ActiveX control in the specified document.
- `ViewPropertyBrowser()`  
  Displays the property window for the selected Microsoft ActiveX control in the specified document.
- `UndoClear()`  
  Clears the list of actions that can be undone for the specified document.
- `PresentIt()`  
  Opens PowerPoint with the specified Word document loaded.
- `SendFax(Address As String, [Subject As Variant])`  
  Sends the specified document as a fax, without any user interaction.
    - `Address As String` (required): The recipient's fax number.
    - `Subject As Variant` (optional): The text for the subject line. The character limit is 255.
- `ClosePrintPreview()`  
  Switches the specified document from print preview to the previous view.
- `CheckConsistency()`  
  Searches all text in a Japanese language document and displays instances where character usage is inconsistent for the same words.
- `CreateLetterContent(DateFormat As String, IncludeHeaderFooter As Boolean, PageDesign As String, LetterStyle As WdLetterStyle, Letterhead As Boolean, LetterheadLocation As WdLetterheadLocation, LetterheadSize As Single, RecipientName As String, RecipientAddress As String, Salutation As String, SalutationType As WdSalutationType, RecipientReference As String, MailingInstructions As String, AttentionLine As String, Subject As String, CCList As String, ReturnAddress As String, SenderName As String, Closing As String, SenderCompany As String, SenderJobTitle As String, SenderInitials As String, EnclosureNumber As Long, [InfoBlock As Variant], [RecipientCode As Variant], [RecipientGender As Variant], [ReturnAddressShortForm As Variant], [SenderCity As Variant], [SenderCode As Variant], [SenderGender As Variant], [SenderReference As Variant]) As LetterContent`  
  Creates and returns a LetterContent object based on the specified letter elements. LetterContent object.
    - `DateFormat As String` (required): The date for the letter.
    - `IncludeHeaderFooter As Boolean` (required): True to include the header and footer from the page design template.
    - `PageDesign As String` (required): The name of the template attached to the document.
    - `LetterStyle As WdLetterStyle` (required): The document layout.
    - `Letterhead As Boolean` (required): True to reserve space for a preprinted letterhead.
    - `LetterheadLocation As WdLetterheadLocation` (required): The location of the preprinted letterhead.
    - `LetterheadSize As Single` (required): The amount of space (in points) to be reserved for a preprinted letterhead.
    - `RecipientName As String` (required): The name of the person who'll be receiving the letter.
    - `RecipientAddress As String` (required): The mailing address of the person who'll be receiving the letter.
    - `Salutation As String` (required): The salutation text for the letter.
    - `SalutationType As WdSalutationType` (required): The salutation type for the letter.
    - `RecipientReference As String` (required): The reference line text for the letter (for example, "In reply to:").
    - `MailingInstructions As String` (required): The mailing instruction text for the letter (for example, "Certified Mail").
    - `AttentionLine As String` (required): The attention line text for the letter (for example, "Attention:").
    - `Subject As String` (required): The subject text for the specified letter.
    - `CCList As String` (required): The names of the carbon copy (CC) recipients for the letter.
    - `ReturnAddress As String` (required): The text of the return mailing address for the letter.
    - `SenderName As String` (required): The name of the person sending the letter.
    - `Closing As String` (required): The closing text for the letter.
    - `SenderCompany As String` (required): The company name of the person creating the letter.
    - `SenderJobTitle As String` (required): The job title of the person creating the letter.
    - `SenderInitials As String` (required): The initials of the person creating the letter.
    - `EnclosureNumber As Long` (required): The number of enclosures for the letter.
    - `InfoBlock As Variant` (optional): This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `RecipientCode As Variant` (optional): This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `RecipientGender As Variant` (optional): This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `ReturnAddressShortForm As Variant` (optional): This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `SenderCity As Variant` (optional): This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `SenderCode As Variant` (optional): This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `SenderGender As Variant` (optional): This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `SenderReference As Variant` (optional): This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
- `AcceptAllRevisions()`  
  Accepts all tracked changes in the specified document.
- `RejectAllRevisions()`  
  Rejects all tracked changes in the specified document.
- `DetectLanguage()`  
  Analyzes the specified text to determine the language that it is written in.
- `ApplyTheme(Name As String)`  
  Applies a theme to an open document.
    - `Name As String` (required): The name of the theme plus any theme formatting options you want to apply. The format of this string is "themennn" where theme and nnn are defined as follows:
- `RemoveTheme()`  
  Removes the active theme from the current document.
- `WebPagePreview()`  
  Displays a preview of the current document as it would look if saved as a webpage.
- `ReloadAs(Encoding As MsoEncoding)`  
  Reloads a document based on an HTML document, using the specified document encoding.
    - `Encoding As MsoEncoding` (required): Specifies the encoding to use when reloading the document.
- `ConvertVietDoc(CodePageOrigin As Long)`  
  Reconverts a Vietnamese document to Unicode using a code page other than the default.
    - `CodePageOrigin As Long` (required): The original code page used to encode the document.
- `PrintOut([Background As Variant], [Append As Variant], [Range As Variant], [OutputFileName As Variant], [From As Variant], [To As Variant], [Item As Variant], [Copies As Variant], [Pages As Variant], [PageType As Variant], [PrintToFile As Variant], [Collate As Variant], [ActivePrinterMacGX As Variant], [ManualDuplexPrint As Variant], [PrintZoomColumn As Variant], [PrintZoomRow As Variant], [PrintZoomPaperWidth As Variant], [PrintZoomPaperHeight As Variant])`  
  Prints all or part of the specified document.
    - `Background As Variant` (optional): Set to True to have the macro continue while Microsoft Word prints the document.
    - `Append As Variant` (optional): Set to True to append the specified document to the file name specified by the OutputFileName argument. False to overwrite the contents of OutputFileName.
    - `Range As Variant` (optional): The page range. Can be any WdPrintOutRange constant.
    - `OutputFileName As Variant` (optional): If PrintToFile is True, this argument specifies the path and file name of the output file.
    - `From As Variant` (optional): The starting page number when Range is set to wdPrintFromTo.
    - `To As Variant` (optional): The ending page number when Range is set to wdPrintFromTo.
    - `Item As Variant` (optional): The item to be printed. Can be any WdPrintOutItem constant.
    - `Copies As Variant` (optional): The number of copies to be printed.
    - `Pages As Variant` (optional): The page numbers and page ranges to be printed, separated by commas. For example, "2, 6-10" prints page 2 and pages 6 through 10.
    - `PageType As Variant` (optional): The type of pages to be printed. Can be any WdPrintOutPages constant.
    - `PrintToFile As Variant` (optional): True to send printer instructions to a file. Make sure to specify a file name with OutputFileName.
    - `Collate As Variant` (optional): When printing multiple copies of a document, True to print all pages of the document before printing the next copy.
    - `ActivePrinterMacGX As Variant` (optional): This argument is available only in Microsoft Office Macintosh Edition. For additional information about this argument, consult the language reference Help included with Microsoft Office Macintosh Edition.
    - `ManualDuplexPrint As Variant` (optional): True to print a two-sided document on a printer without a duplex printing kit. If this argument is True, the PrintBackground and PrintReverse properties are ignored. Use the PrintOddPagesInAscendingOrder and PrintEvenPagesInAscendingOrder properties to control the output during manual duplex printing. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `PrintZoomColumn As Variant` (optional): The number of pages you want Word to fit horizontally on one page. Can be 1, 2, 3, or 4. Use with the PrintZoomRow argument to print multiple pages on a single sheet.
    - `PrintZoomRow As Variant` (optional): The number of pages you want Word to fit vertically on one page. Can be 1, 2, or 4. Use with the PrintZoomColumn argument to print multiple pages on a single sheet.
    - `PrintZoomPaperWidth As Variant` (optional): The width to which you want Word to scale printed pages, in twips (20 twips = 1 point; 72 points = 1 inch).
    - `PrintZoomPaperHeight As Variant` (optional): The height to which you want Word to scale printed pages, in twips (20 twips = 1 point; 72 points = 1 inch).
- `CheckIn([SaveChanges As Boolean], [Comments As Variant], [MakePublic As Boolean])`  
  Returns a document from a local computer to a server, and sets the local document to read-only so that it cannot be edited locally.
    - `SaveChanges As Boolean` (optional): True saves the document to the server location. The default is True.
    - `Comments As Variant` (optional): Comments for the revision of the document being checked in (only applies if SaveChanges equals True).
    - `MakePublic As Boolean` (optional): True allows the user to perform a publish on the document after being checked in. This submits the document for the approval process, which can eventually result in a version of the document being published to users with read-only rights to the document (only applies if SaveChanges equals True). The default is False.
- `CanCheckin() As Boolean`  
  True if Microsoft Word can check in a specified document to a server. Read/write Boolean.
- `Merge(FileName As String, [MergeTarget As Variant], [DetectFormatChanges As Variant], [UseFormattingFrom As Variant], [AddToRecentFiles As Variant])`  
  Merges the changes marked with revision marks from one document to another.
    - `MergeTarget As Variant` (optional): Specifies where to place the final merged content.
    - `DetectFormatChanges As Variant` (optional): Specifies whether or not to mark formatting differences.
    - `UseFormattingFrom As Variant` (optional): Specifies which document to use for formatting in the merged document.
    - `AddToRecentFiles As Variant` (optional): Specifies whether to add the document in the Name parameter to the list of recent files.
- `SendForReview([Recipients As Variant], [Subject As Variant], [ShowMessage As Variant], [IncludeAttachment As Variant])`  
  Sends a document in an email message for review by the specified recipients.
    - `Recipients As Variant` (optional): A string that lists the people to whom to send the message. These can be unresolved names and aliases in an email phone book or full email addresses. Separate multiple recipients with a semicolon (;). If left blank and ShowMessage is False, you will receive an error message and the message will not be sent.
    - `Subject As Variant` (optional): A string for the subject of the message. If left blank, the subject will be: Please review "file name".
    - `ShowMessage As Variant` (optional): A Boolean value that indicates whether the message should be displayed when the method is executed. The default value is True. If set to False, the message is automatically sent to the recipients without first showing the message to the sender.
    - `IncludeAttachment As Variant` (optional): A Boolean value that indicates whether the message should include an attachment or a link to a server location. The default value is True. If set to False, the document must be stored at a shared location.
- `ReplyWithChanges([ShowMessage As Variant])`  
  Sends an email message to the author of a document that has been sent out for review, notifying them that a reviewer has completed review of the document.
    - `ShowMessage As Variant` (optional): True to display the message prior to sending. False to automatically send the message without displaying it first. The default value is True.
- `EndReview()`  
  Terminates a review of a file that has been sent for review using the SendForReview method or that has been automatically placed in a review cycle by sending a document to another user in an email message.
- `SetPasswordEncryptionOptions(PasswordEncryptionProvider As String, PasswordEncryptionAlgorithm As String, PasswordEncryptionKeyLength As Long, [PasswordEncryptionFileProperties As Variant])`  
  Sets the options Microsoft Word uses for encrypting documents with passwords.
    - `PasswordEncryptionProvider As String` (required): The name of the encryption provider.
    - `PasswordEncryptionAlgorithm As String` (required): The name of the encryption algorithm. Word supports stream-encrypted algorithms.
    - `PasswordEncryptionKeyLength As Long` (required): The encryption key length. Must be a multiple of 8, starting at 40.
    - `PasswordEncryptionFileProperties As Variant` (optional): True for Word to encrypt file properties. Default is True.
- `SetDefaultTableStyle(Style As Variant, SetInTemplate As Boolean)`  
  Specifies the table style to use for newly created tables in a document.
    - `Style As Variant` (required): A string specifying the name of the style.
    - `SetInTemplate As Boolean` (required): True to save the table style in the template attached to the document.
- `DeleteAllComments()`  
  Deletes all comments from the Comments collection in a document.
- `AcceptAllRevisionsShown()`  
  Accepts all revisions in the specified document that are displayed on the screen.
- `RejectAllRevisionsShown()`  
  Rejects all revisions in a document that are displayed on the screen.
- `DeleteAllCommentsShown()`  
  Deletes all revisions in a specified document that are displayed on the screen.
- `ResetFormFields()`  
  Clears all form fields in a document, preparing the form to be filled in again.
- `SendFaxOverInternet([Recipients As Variant], [Subject As Variant], [ShowMessage As Variant])`  
  Sends a document to a fax service provider, who faxes the document to one or more specified recipients.
    - `Recipients As Variant` (optional): A String that represents the fax numbers and email addresses of the people to whom to send the fax. Separate multiple recipients with a semicolon.
    - `Subject As Variant` (optional): A String that represents the subject line for the faxed document.
    - `ShowMessage As Variant` (optional): True displays the fax message before sending it. False sends the fax without displaying the fax message.
- `TransformDocument(Path As String, [DataOnly As Boolean])`  
  Applies the specified Extensible Stylesheet Language Transformation (XSLT) file to the specified document and replaces the document with the results.
    - `Path As String` (required): The path for the XSLT to use.
    - `DataOnly As Boolean` (optional): True applies the transformation only to the data in the document, excluding Microsoft Word XML. False applies the transform to the entire document, including Word XML. Default value is True.
- `Protect(Type As WdProtectionType, [NoReset As Variant], [Password As Variant], [UseIRM As Variant], [EnforceStyleLock As Variant])`  
  Protects the specified document from unauthorized changes.
- `SelectAllEditableRanges([EditorID As Variant])`  
  Selects all ranges for which the specified user or group of users has permission to modify.
    - `EditorID As Variant` (optional): Can be either a String that represents the user's email alias (if in the same domain), an email address, or a WdEditorType constant that represents a group of users. If omitted, only ranges for which all users have permissions will be selected.
- `DeleteAllEditableRanges([EditorID As Variant])`  
  Deletes permissions in all ranges for which the specified user or group of users has permission to modify.
    - `EditorID As Variant` (optional): Can be either a String that represents the user's email alias (if in the same domain), an email address, or a WdEditorType constant that represents a group of users. If omitted, no permissions are deleted from a document.
- `DeleteAllInkAnnotations()`  
  Deletes all handwritten ink annotations in a document.
- `Compare(Name As String, [AuthorName As Variant], [CompareTarget As Variant], [DetectFormatChanges As Variant], [IgnoreAllComparisonWarnings As Variant], [AddToRecentFiles As Variant], [RemovePersonalInformation As Variant], [RemoveDateAndTime As Variant])`  
  Displays revision marks that indicate where the specified document differs from another document.
    - `Name As String` (required): The name of the document with which the specified document is compared.
    - `AuthorName As Variant` (optional): The reviewer name associated with the differences generated by the comparison. If unspecified, the value defaults to the author name of the revised document or the string "Comparison" if no author information is present.
    - `CompareTarget As Variant` (optional): The target document for the comparison. Can be any WdCompareTarget constant.
    - `DetectFormatChanges As Variant` (optional): True (default) for the comparison to include detection of format changes.
    - `IgnoreAllComparisonWarnings As Variant` (optional): True compares the documents without notifying a user of problems. The default value is False.
    - `AddToRecentFiles As Variant` (optional): True adds the document to the list of recently used files on the File menu.
    - `RemovePersonalInformation As Variant` (optional): True removes all user information from comments, revisions, and the properties dialog box in the returned Document object . The default value is False.
    - `RemoveDateAndTime As Variant` (optional): True removes date and time stamp information from tracked changes in the returned Document object. The default value is False.
- `RemoveLockedStyles()`  
  Purges a document of locked styles when formatting restrictions have been applied in a document.
- `SelectSingleNode(XPath As String, [PrefixMapping As String], [FastSearchSkippingTextNodes As Boolean]) As XMLNode`  
  Returns an XMLNode object that represents the first node that matches the XPath parameter in the specified document.
    - `XPath As String` (required): A valid XPath string. For more information on XPath, see the XPath reference documentation on the Microsoft Developer Network (MSDN) Web site.
    - `PrefixMapping As String` (optional): Provides the prefix in the schema against which to perform the search. Use the PrefixMapping parameter if your XPath parameter uses names to search for elements.
    - `FastSearchSkippingTextNodes As Boolean` (optional): True skips all text nodes while searching for the specified node. False includes text nodes in the search. Default value is True.
- `SelectNodes(XPath As String, [PrefixMapping As String], [FastSearchSkippingTextNodes As Boolean]) As XMLNodes`  
  Returns an XMLNodes collection that represents all the nodes that match the XPath parameter in the order in which they appear in the document or range.
    - `XPath As String` (required): A valid XPath string. For more information on XPath, see the XPath reference documentation on the Microsoft Developer Network (MSDN) Web site.
    - `PrefixMapping As String` (optional): Provides the prefix in the schema against which to perform the search. Use the PrefixMapping parameter if your XPath parameter uses names to search for elements.
    - `FastSearchSkippingTextNodes As Boolean` (optional): True skips all text nodes while searching for the specified node. False includes text nodes in the search. Default value is False.
- `RemoveDocumentInformation(RemoveDocInfoType As WdRemoveDocInfoType)`  
  Removes sensitive information, properties, comments, and other metadata from a document.
    - `RemoveDocInfoType As WdRemoveDocInfoType` (required): Specifies what to remove.
- `CheckInWithVersion([SaveChanges As Boolean], [Comments As Variant], [MakePublic As Boolean], [VersionType As Variant])`  
  Saves a document to a server from a local computer, and sets the local document to read-only so that it cannot be edited locally.
    - `SaveChanges As Boolean` (optional): True to save the document to the server location. The default is True.
    - `Comments As Variant` (optional): Comments for the revision of the document being checked in (applies only if SaveChanges is set to True).
    - `MakePublic As Boolean` (optional): True to allow the user to publish the document after it is checked in.
    - `VersionType As Variant` (optional): Specifies versioning information for the document.
- `LockServerFile()`  
  Locks the file on the server preventing anyone else from editing it.
- `GetWorkflowTasks() As WorkflowTasks`  
  Returns a WorkflowTasks collection that represents the workflow tasks assigned to a document.
- `GetWorkflowTemplates() As WorkflowTemplates`  
  Returns a WorkflowTemplates collection that represents the workflow templates attached to a document.
- `SaveAsQuickStyleSet(FileName As String)`  
  Saves the group of quick styles currently in use.
    - `FileName As String` (required): The path and file name for the quick style set file.
- `ApplyDocumentTheme(FileName As String)`
- `SelectLinkedControls(Node As CustomXMLNode) As ContentControls`  
  Returns a ContentControls collection that represents all content controls in a document that are linked to the specific custom XML node in the document's XML data store as specified by the Node parameter. Read-only.
    - `Node As CustomXMLNode` (required): The XML node in the document's data store to which the content controls are linked.
- `SelectUnlinkedControls([Stream As CustomXMLPart]) As ContentControls`  
  Returns a ContentControls collection that represents all content controls in a document that are not linked to an XML node in the document's XML data store. Read-only.
    - `Stream As CustomXMLPart` (optional): A custom XML part reference. Setting this parameter filters the returned content controls to include only content controls that reference this CustomXMLPart in their XMLMapping definition.
- `SelectContentControlsByTitle(Title As String) As ContentControls`  
  Returns a ContentControls collection that represents all the content controls in a document with the title specified in the Title parameter. Read-only.
- `ExportAsFixedFormat(OutputFileName As String, ExportFormat As WdExportFormat, [OpenAfterExport As Boolean], [OptimizeFor As WdExportOptimizeFor], [Range As WdExportRange], [From As Long], [To As Long], [Item As WdExportItem], [IncludeDocProps As Boolean], [KeepIRM As Boolean], [CreateBookmarks As WdExportCreateBookmarks], [DocStructureTags As Boolean], [BitmapMissingFonts As Boolean], [UseISO19005_1 As Boolean], [FixedFormatExtClassPtr As Variant])`  
  Saves a document in PDF or XPS format.
    - `OutputFileName As String` (required): The path and file name of the new PDF or XPS file.
    - `ExportFormat As WdExportFormat` (required): Specifies either PDF or XPS format.
    - `OpenAfterExport As Boolean` (optional): Opens the new file after exporting the contents.
    - `OptimizeFor As WdExportOptimizeFor` (optional): Specifies whether to optimize for screen or print.
    - `Range As WdExportRange` (optional): Specifies whether the export range is the entire document, the current page, a range of text, or the current selection. The default is to export the entire document.
    - `From As Long` (optional): Specifies the starting page number, if the Range parameter is set to wdExportFromTo.
    - `To As Long` (optional): Specifies the ending page number, if the Range parameter is set to wdExportFromTo.
    - `Item As WdExportItem` (optional): Specifies whether the export process includes text only or includes text with markup.
    - `IncludeDocProps As Boolean` (optional): Specifies whether to include document properties in the newly exported file.
    - `KeepIRM As Boolean` (optional): Specifies whether to copy IRM permissions to an XPS document if the source document has IRM protections.</br></br>If ExportFormat is wdExportFormatPDF, this flag also specifies whether to copy labels to the PDF.</br></br>Default value is True.
    - `CreateBookmarks As WdExportCreateBookmarks` (optional): Specifies whether to export bookmarks and the type of bookmarks to export.
    - `DocStructureTags As Boolean` (optional): Specifies whether to include extra data to help screen readers, for example information about the flow and logical organization of the content. Default value is True.
    - `BitmapMissingFonts As Boolean` (optional): Specifies whether to include a bitmap of the text. Set this parameter to True when font licenses don't permit a font to be embedded in the PDF file. If False, the font is referenced, and the viewer's computer substitutes an appropriate font if the authored one is not available. Default value is True.
    - `FixedFormatExtClassPtr As Variant` (optional): Specifies a pointer to an add-in that allows calls to an alternate implementation of code. The alternate implementation of code interprets the EMF and EMF+ page descriptions that are generated by the applications to make their own PDF or XPS. For more information, see Extend the fixed-format export feature in Word Automation Services.
- `FreezeLayout()`  
  In Web view, fixes the layout of the document as it currently appears so that line breaks remain fixed and ink annotations don't move when you resize the window.
- `DowngradeDocument()`  
  Downgrades a document to the Word 97-2003 document format so that it can be edited in a previous version of Microsoft Word. .
- `Convert()`  
  Converts file to the newest file format and enables all new features.
- `SelectContentControlsByTag(Tag As String) As ContentControls`  
  Returns a ContentControls collection that represents all the content controls in a document with the tag value specified in the Tag parameter. Read-only.
    - `Tag As String` (required): The tag value of the content controls to return.
- `ConvertAutoHyphens()`  
  Converts hyphens created by automatic hyphenation to manual hyphens.
- `ApplyQuickStyleSet2(Style As Variant)`  
  Applies the specified Quick Style set to the document.
    - `Style As Variant` (required): Can be either a String that specifies the name of the set to use (corresponds to the name listed in the Style Set list) or a constant from the WdApplyQuickStyleSets enumeration.
- `SaveAs2([FileName As Variant], [FileFormat As Variant], [LockComments As Variant], [Password As Variant], [AddToRecentFiles As Variant], [WritePassword As Variant], [ReadOnlyRecommended As Variant], [EmbedTrueTypeFonts As Variant], [SaveNativePictureFormat As Variant], [SaveFormsData As Variant], [SaveAsAOCELetter As Variant], [Encoding As Variant], [InsertLineBreaks As Variant], [AllowSubstitutions As Variant], [LineEnding As Variant], [AddBiDiMarks As Variant], [CompatibilityMode As Variant])`
- `SetCompatibilityMode(Mode As Long)`
- `ReturnToLastReadPosition() As Long`  
  Returns the active document to the last saved reading position, and returns an integer that represents that position.
- `ExportAsFixedFormat2(OutputFileName As String, ExportFormat As WdExportFormat, [OpenAfterExport As Boolean], [OptimizeFor As WdExportOptimizeFor], [Range As WdExportRange], [From As Long], [To As Long], [Item As WdExportItem], [IncludeDocProps As Boolean], [KeepIRM As Boolean], [CreateBookmarks As WdExportCreateBookmarks], [DocStructureTags As Boolean], [BitmapMissingFonts As Boolean], [UseISO19005_1 As Boolean], [OptimizeForImageQuality As Boolean], [FixedFormatExtClassPtr As Variant])`  
  Saves a document in PDF or XPS format.
    - `OutputFileName As String` (required): The path and file name of the new PDF or XPS file.
    - `ExportFormat As WdExportFormat` (required): Specifies either PDF or XPS format.
    - `OpenAfterExport As Boolean` (optional): Opens the new file after exporting the contents.
    - `OptimizeFor As WdExportOptimizeFor` (optional): Specifies whether to optimize for screen or print.
    - `Range As WdExportRange` (optional): Specifies whether the export range is the entire document, the current page, a range of text, or the current selection. The default is to export the entire document.
    - `From As Long` (optional): Specifies the starting page number, if the Range parameter is set to wdExportFromTo.
    - `To As Long` (optional): Specifies the ending page number, if the Range parameter is set to wdExportFromTo.
    - `Item As WdExportItem` (optional): Specifies whether the export process includes text only or includes text with markup.
    - `IncludeDocProps As Boolean` (optional): Specifies whether to include document properties in the newly exported file.
    - `KeepIRM As Boolean` (optional): Specifies whether to copy IRM permissions to an XPS document if the source document has IRM protections.</br></br>If ExportFormat is wdExportFormatPDF, this flag also specifies whether to copy labels to the PDF.</br></br>Default value is True.
    - `CreateBookmarks As WdExportCreateBookmarks` (optional): Specifies whether to export bookmarks and the type of bookmarks to export.
    - `DocStructureTags As Boolean` (optional): Specifies whether to include extra data to help screen readers, for example information about the flow and logical organization of the content. Default value is True.
    - `BitmapMissingFonts As Boolean` (optional): Specifies whether to include a bitmap of the text. Set this parameter to True when font licenses don't permit a font to be embedded in the PDF file. If False, the font is referenced, and the viewer's computer substitutes an appropriate font if the authored one is not available. Default value is True.
    - `OptimizeForImageQuality As Boolean` (optional): Specifies whether to downsample images or keep their original quality. If True, the resulting files will have better image quality but may be larger. Default value is False.
    - `FixedFormatExtClassPtr As Variant` (optional): Specifies a pointer to an add-in that allows calls to an alternate implementation of code. The alternate implementation of code interprets the EMF and EMF+ page descriptions that are generated by the applications to make their own PDF or XPS. For more information, see Extend the fixed-format export feature in Word Automation Services.
- `ExportAsFixedFormat3(OutputFileName As String, ExportFormat As WdExportFormat, [OpenAfterExport As Boolean], [OptimizeFor As WdExportOptimizeFor], [Range As WdExportRange], [From As Long], [To As Long], [Item As WdExportItem], [IncludeDocProps As Boolean], [KeepIRM As Boolean], [CreateBookmarks As WdExportCreateBookmarks], [DocStructureTags As Boolean], [BitmapMissingFonts As Boolean], [UseISO19005_1 As Boolean], [OptimizeForImageQuality As Boolean], [ImproveExportTagging As Boolean], [FixedFormatExtClassPtr As Variant])`  
  Saves a document in PDF or XPS format.
    - `OutputFileName As String` (required): The path and file name of the new PDF or XPS file.
    - `ExportFormat As WdExportFormat` (required): Specifies either PDF or XPS format.
    - `OpenAfterExport As Boolean` (optional): Opens the new file after exporting the contents.
    - `OptimizeFor As WdExportOptimizeFor` (optional): Specifies whether to optimize for screen or print.
    - `Range As WdExportRange` (optional): Specifies whether the export range is the entire document, the current page, a range of text, or the current selection. The default is to export the entire document.
    - `From As Long` (optional): Specifies the starting page number, if the Range parameter is set to wdExportFromTo.
    - `To As Long` (optional): Specifies the ending page number, if the Range parameter is set to wdExportFromTo.
    - `Item As WdExportItem` (optional): Specifies whether the export process includes text only or includes text with markup.
    - `IncludeDocProps As Boolean` (optional): Specifies whether to include document properties in the newly exported file.
    - `KeepIRM As Boolean` (optional): Specifies whether to copy IRM permissions to an XPS document if the source document has IRM protections.</br></br>If ExportFormat is wdExportFormatPDF, this flag also specifies whether to copy labels to the PDF.</br></br>Default value is True.
    - `CreateBookmarks As WdExportCreateBookmarks` (optional): Specifies whether to export bookmarks and the type of bookmarks to export.
    - `DocStructureTags As Boolean` (optional): Specifies whether to include extra data to help screen readers, for example information about the flow and logical organization of the content. Default value is True.
    - `BitmapMissingFonts As Boolean` (optional): Specifies whether to include a bitmap of the text. Set this parameter to True when font licenses don't permit a font to be embedded in the PDF file. If False, the font is referenced, and the viewer's computer substitutes an appropriate font if the authored one is not available. Default value is True.
    - `OptimizeForImageQuality As Boolean` (optional): Specifies whether to downsample images or keep their original quality. If True, the resulting files will have better image quality but may be larger. Default value is False.
    - `ImproveExportTagging As Boolean` (optional): Specifies whether to enable improved accessbility tagging. For more information, see the Remarks section. Default value is False.
    - `FixedFormatExtClassPtr As Variant` (optional): Specifies a pointer to an add-in that allows calls to an alternate implementation of code. The alternate implementation of code interprets the EMF and EMF+ page descriptions that are generated by the applications to make their own PDF or XPS.

## Events (14)

- `New()`  
  Occurs when a new document based on the template is created. A procedure for the New event will run only if it is stored in a template.
- `Open()`  
  Occurs when a document is opened.
- `Close()`  
  Occurs when a document is closed.
- `Sync(SyncEventType As MsoSyncEventType)`
    - `SyncEventType As MsoSyncEventType` (required): The status of the document synchronization.
- `XMLAfterInsert(NewXMLNode As XMLNode, InUndoRedo As Boolean)`  
  Occurs when a user adds a new XML element to a document. If more than one element is added to the document at the same time (for example, when cutting and pasting XML), the event fires for each element that is inserted.
    - `NewXMLNode As XMLNode` (required): The newly added XML node.
    - `InUndoRedo As Boolean` (required): True indicates the action was performed using the Undo or Redo feature in Microsoft Word.
- `XMLBeforeDelete(DeletedRange As Range, OldXMLNode As XMLNode, InUndoRedo As Boolean)`  
  Occurs when a user deletes an XML element from a document. If more than one element is deleted from the document at the same time (for example, when cutting and pasting XML), the event fires for each element that is deleted.
    - `DeletedRange As Range` (required): The contents of the XML element being deleted. If only an element is deleted and not associated text, the DeletedRange parameter will not exist and will, therefore, be set to Nothing.
    - `OldXMLNode As XMLNode` (required): The node that is being deleted.
    - `InUndoRedo As Boolean` (required): True indicates the action was performed using the Undo or Redo feature in Microsoft Word.
- `ContentControlAfterAdd(NewContentControl As ContentControl, InUndoRedo As Boolean)`  
  Occurs after adding a content control to a document.
    - `NewContentControl As ContentControl` (required): The content control being added.
    - `InUndoRedo As Boolean` (required): Specifies whether the addition is taking place as part an undo or redo action.
- `ContentControlBeforeDelete(OldContentControl As ContentControl, InUndoRedo As Boolean)`  
  Occurs before removing a content control from a document.
    - `OldContentControl As ContentControl` (required): The content control being deleted.
    - `InUndoRedo As Boolean` (required): Specifies whether the removal is taking place as part an undo or redo action.
- `ContentControlOnExit(ContentControl As ContentControl, Cancel As Boolean)`  
  Occurs when a user leaves a content control.
    - `ContentControl As ContentControl` (required): The content control that the user is leaving.
    - `Cancel As Boolean` (required): Specifies whether to cancel the event. True cancels the event and does not allow the user to leave the control.
- `ContentControlOnEnter(ContentControl As ContentControl)`  
  Occurs when a user enters a content control.
    - `ContentControl As ContentControl` (required): The content control that the user is entering.
- `ContentControlBeforeStoreUpdate(ContentControl As ContentControl, Content As String)`  
  Occurs before updating the document's XML data store with the value of a content control.
    - `ContentControl As ContentControl` (required): The content control being updated.
    - `Content As String` (required): The content being stored for a control in the document data store. Use this parameter to change the XML data before sending the value to the XML data store.
- `ContentControlBeforeContentUpdate(ContentControl As ContentControl, Content As String)`  
  Occurs before updating the content in a content control, but only when the content comes from the Office XML data store.
    - `ContentControl As ContentControl` (required): The content control being updated.
    - `Content As String` (required): The updated content for a control. Use this parameter to change the contents of the XML data and format it for display.
- `BuildingBlockInsert(Range As Range, Name As String, Category As String, BlockType As String, Template As String)`  
  Occurs when you insert a building block into a document. .
    - `Range As Range` (required): Specifies the position where the building block is inserted.
    - `Name As String` (required): Specifies the name of the building block.
    - `Category As String` (required): Specifies the building block category.
    - `Template As String` (required): Specifies the name of the template that contains the building block.
- `ContentControlNonContentChange(ContentControl As ContentControl)`
