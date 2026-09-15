# WdWordDialog

**Type:** Enumeration  
**Library:** Microsoft Word 16.0 Object Library  

Indicates the Microsoft Word dialog boxes with which you can work and specifies arguments, if applicable, that you can use to get or set values in a dialog box.

## Constants (238)

- `wdDialogHelpAbout` = 9  
  APPNAME, APPCOPYRIGHT, APPUSERNAME, APPORGANIZATION, APPSERIALNUMBER
- `wdDialogHelpWordPerfectHelp` = 10  
  WPCommand, HelpText, DemoGuidance
- `wdDialogDocumentStatistics` = 78  
  FileName, Directory, Template, Title, Created, LastSaved, LastSavedBy, Revision, Time, Printed, Pages, Words, Characters, Paragraphs, Lines, FileSize
- `wdDialogFileNew` = 79  
  Template, NewTemplate, DocumentType, Visible
- `wdDialogFileOpen` = 80  
  Name, ConfirmConversions, ReadOnly, LinkToSource, AddToMru, PasswordDoc, PasswordDot, Revert, WritePasswordDoc, WritePasswordDot, Connection, SQLStatement, SQLStatement1, Format, Encoding, Visible, OpenExclusive, OpenAndRepair, SubType, DocumentDirection, NoEncodingDialog, XMLTransform
- `wdDialogMailMergeOpenDataSource` = 81  
  (none)
- `wdDialogMailMergeOpenHeaderSource` = 82  
  (none)
- `wdDialogFileSaveAs` = 84  
  Name, Format, LockAnnot, Password, AddToMru, WritePassword, RecommendReadOnly, EmbedFonts, NativePictureFormat, FormsData, SaveAsAOCELetter, WriteVersion, VersionDesc, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks
- `wdDialogFileSummaryInfo` = 86  
  Title, Subject, Author, Keywords, Comments, FileName, Directory, Template, CreateDate, LastSavedDate, LastSavedBy, RevisionNumber, EditTime, LastPrintedDate, NumPages, NumWords, NumChars, NumParas, NumLines, Update, FileSize
- `wdDialogToolsTemplates` = 87  
  (none)
- `wdDialogFilePrint` = 88  
  Background, AppendPrFile, Range, PrToFileName, From, To, Type, NumCopies, Pages, Order, PrintToFile, Collate, FileName, Printer, OutputPrinter, DuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight, ZoomPaper
- `wdDialogFilePrintSetup` = 97  
  Printer, Options, Network, DoNotSetAsSysDefault
- `wdDialogFileFind` = 99  
  SearchName, SearchPath, Name, SubDir, Title, Author, Keywords, Subject, Options, MatchCase, Text, PatternMatch, DateSavedFrom, DateSavedTo, SavedBy, DateCreatedFrom, DateCreatedTo, View, SortBy, ListBy, SelectedFile, Add, Delete, ShowFolders, MatchByte
- `wdDialogFormatAddrFonts` = 103  
  Points, Underline, Color, StrikeThrough, Superscript, Subscript, Hidden, SmallCaps, AllCaps, Spacing, Position, Kerning, KerningMin, Default, Tab, Font, Bold, Italic, DoubleStrikeThrough, Shadow, Outline, Emboss, Engrave, Scale, Animations, CharAccent, FontMajor, FontLowAnsi, FontHighAnsi, CharacterWidthGrid, ColorRGB, UnderlineColor, PointsBi, ColorBi, FontNameBi, BoldBi, ItalicBi, DiacColor
- `wdDialogEditPasteSpecial` = 111  
  IconNumber, Link, DisplayIcon, Class, DataType, IconFileName, Caption, Floating
- `wdDialogEditFind` = 112  
  Find, Replace, Direction, MatchCase, WholeWord, PatternMatch, SoundsLike, FindNext, ReplaceOne, ReplaceAll, Format, Wrap, FindAllWordForms, MatchByte, FuzzyFind, Destination, CorrectEnd, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl
- `wdDialogEditReplace` = 117  
  Find, Replace, Direction, MatchCase, WholeWord, PatternMatch, SoundsLike, FindNext, ReplaceOne, ReplaceAll, Format, Wrap, FindAllWordForms, MatchByte, FuzzyFind, Destination, CorrectEnd, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl
- `wdDialogEditStyle` = 120  
  (none)
- `wdDialogEditLinks` = 124  
  UpdateMode, Locked, SavePictureInDoc, UpdateNow, OpenSource, KillLink, Link, Application, Item, FileName, PreserveFormatLinkUpdate
- `wdDialogEditObject` = 125  
  Verb
- `wdDialogTableToText` = 128  
  (none)
- `wdDialogTextToTable` = 127  
  (none)
- `wdDialogTableInsertTable` = 129  
  (none)
- `wdDialogTableInsertCells` = 130  
  (none)
- `wdDialogTableInsertRow` = 131  
  (none)
- `wdDialogTableDeleteCells` = 133  
  (none)
- `wdDialogTableSplitCells` = 137  
  (none)
- `wdDialogTableRowHeight` = 142  
  (none)
- `wdDialogTableColumnWidth` = 143  
  (none)
- `wdDialogToolsCustomize` = 152  
  (none)
- `wdDialogInsertBreak` = 159  
  Type
- `wdDialogInsertSymbol` = 162  
  Font, Tab, CharNum, CharNumLow, Unicode, Hint
- `wdDialogInsertPicture` = 163  
  Name, LinkToFile, New, FloatOverText
- `wdDialogInsertFile` = 164  
  Name, Range, ConfirmConversions, Link, Attachment
- `wdDialogInsertDateTime` = 165  
  DateTimePic, InsertAsField, DbCharField, DateLanguage, CalendarType
- `wdDialogInsertField` = 166  
  Field
- `wdDialogInsertMergeField` = 167  
  MergeField, WordField
- `wdDialogInsertBookmark` = 168  
  Name, SortBy, Add, Delete, Goto, Hidden
- `wdDialogMarkIndexEntry` = 169  
  (none)
- `wdDialogInsertIndex` = 170  
  Outline, Fields, From, To, TableId, AddedStyles, Caption, HeadingSeparator, Replace, MarkEntry, AutoMark, MarkCitation, Type, RightAlignPageNumbers, Passim, KeepFormatting, Columns, Category, Label, ShowPageNumbers, AccentedLetters, Filter, SortBy, Leader, TOCUseHyperlinks, TOCHidePageNumInWeb, IndexLanguage, UseOutlineLevel
- `wdDialogInsertTableOfContents` = 171  
  Outline, Fields, From, To, TableId, AddedStyles, Caption, HeadingSeparator, Replace, MarkEntry, AutoMark, MarkCitation, Type, RightAlignPageNumbers, Passim, KeepFormatting, Columns, Category, Label, ShowPageNumbers, AccentedLetters, Filter, SortBy, Leader, TOCUseHyperlinks, TOCHidePageNumInWeb, IndexLanguage, UseOutlineLevel
- `wdDialogInsertObject` = 172  
  IconNumber, FileName, Link, DisplayIcon, Tab, Class, IconFileName, Caption, Floating
- `wdDialogToolsCreateEnvelope` = 173  
  (none)
- `wdDialogFormatFont` = 174  
  Points, Underline, Color, StrikeThrough, Superscript, Subscript, Hidden, SmallCaps, AllCaps, Spacing, Position, Kerning, KerningMin, Default, Tab, Font, Bold, Italic, DoubleStrikeThrough, Shadow, Outline, Emboss, Engrave, Scale, Animations, CharAccent, FontMajor, FontLowAnsi, FontHighAnsi, CharacterWidthGrid, ColorRGB, UnderlineColor, PointsBi, ColorBi, FontNameBi, BoldBi, ItalicBi, DiacColor
- `wdDialogFormatParagraph` = 175  
  LeftIndent, RightIndent, Before, After, LineSpacingRule, LineSpacing, Alignment, WidowControl, KeepWithNext, KeepTogether, PageBreak, NoLineNum, DontHyphen, Tab, FirstIndent, OutlineLevel, Kinsoku, WordWrap, OverflowPunct, TopLinePunct, AutoSpaceDE, LineHeightGrid, AutoSpaceDN, CharAlign, CharacterUnitLeftIndent, AdjustRight, CharacterUnitFirstIndent, CharacterUnitRightIndent, LineUnitBefore, LineUnitAfter, NoSpaceBetweenParagraphsOfSameStyle, OrientationBi
- `wdDialogFormatSectionLayout` = 176  
  SectionStart, VertAlign, Endnotes, LineNum, StartingNum, FromText, CountBy, NumMode, SectionType
- `wdDialogFormatColumns` = 177  
  Columns, ColumnNo, ColumnWidth, ColumnSpacing, EvenlySpaced, ApplyColsTo, ColLine, StartNewCol, FlowColumnsRtl
- `wdDialogFileDocumentLayout` = 178  
  Tab, PaperSize, TopMargin, BottomMargin, LeftMargin, RightMargin, Gutter, PageWidth, PageHeight, Orientation, FirstPage, OtherPages, VertAlign, ApplyPropsTo, Default, FacingPages, HeaderDistance, FooterDistance, SectionStart, OddAndEvenPages, DifferentFirstPage, Endnotes, LineNum, StartingNum, FromText, CountBy, NumMode, TwoOnOne, GutterPosition, LayoutMode, CharsLine, LinesPage, CharPitch, LinePitch, DocFontName, DocFontSize, PageColumns, TextFlow, FirstPageOnLeft, SectionType, RTLAlignment
- `wdDialogFilePageSetup` = 178  
  Tab, PaperSize, TopMargin, BottomMargin, LeftMargin, RightMargin, Gutter, PageWidth, PageHeight, Orientation, FirstPage, OtherPages, VertAlign, ApplyPropsTo, Default, FacingPages, HeaderDistance, FooterDistance, SectionStart, OddAndEvenPages, DifferentFirstPage, Endnotes, LineNum, StartingNum, FromText, CountBy, NumMode, TwoOnOne, GutterPosition, LayoutMode, CharsLine, LinesPage, CharPitch, LinePitch, DocFontName, DocFontSize, PageColumns, TextFlow, FirstPageOnLeft, SectionType, RTLAlignment, FolioPrint
- `wdDialogFormatTabs` = 179  
  Position, DefTabs, Align, Leader, Set, Clear, ClearAll
- `wdDialogFormatStyle` = 180  
  Name, Delete, Merge, NewName, BasedOn, NextStyle, Type, FileName, Source, AddToTemplate, Define, Rename, Apply, New, Link
- `wdDialogFormatDefineStyleFont` = 181  
  Points, Underline, Color, StrikeThrough, Superscript, Subscript, Hidden, SmallCaps, AllCaps, Spacing, Position, Kerning, KerningMin, Default, Tab, Font, Bold, Italic, DoubleStrikeThrough, Shadow, Outline, Emboss, Engrave, Scale, Animations, CharAccent, FontMajor, FontLowAnsi, FontHighAnsi, CharacterWidthGrid, ColorRGB, UnderlineColor, PointsBi, ColorBi, FontNameBi, BoldBi, ItalicBi, DiacColor
- `wdDialogFormatDefineStylePara` = 182  
  LeftIndent, RightIndent, Before, After, LineSpacingRule, LineSpacing, Alignment, WidowControl, KeepWithNext, KeepTogether, PageBreak, NoLineNum, DontHyphen, Tab, FirstIndent, OutlineLevel, Kinsoku, WordWrap, OverflowPunct, TopLinePunct, AutoSpaceDE, LineHeightGrid, AutoSpaceDN, CharAlign, CharacterUnitLeftIndent, AdjustRight, CharacterUnitFirstIndent, CharacterUnitRightIndent, LineUnitBefore, LineUnitAfter, NoSpaceBetweenParagraphsOfSameStyle, OrientationBi
- `wdDialogFormatDefineStyleTabs` = 183  
  Position, DefTabs, Align, Leader, Set, Clear, ClearAll
- `wdDialogFormatDefineStyleFrame` = 184  
  Wrap, WidthRule, FixedWidth, HeightRule, FixedHeight, PositionHorz, PositionHorzRel, DistFromText, PositionVert, PositionVertRel, DistVertFromText, MoveWithText, LockAnchor, RemoveFrame
- `wdDialogFormatDefineStyleBorders` = 185  
  ApplyTo, Shadow, TopBorder, LeftBorder, BottomBorder, RightBorder, HorizBorder, VertBorder, TopColor, LeftColor, BottomColor, RightColor, HorizColor, VertColor, FromText, Shading, Foreground, Background, Tab, FineShading, TopStyle, LeftStyle, BottomStyle, RightStyle, HorizStyle, VertStyle, TopWeight, LeftWeight, BottomWeight, RightWeight, HorizWeight, VertWeight, BorderObjectType, BorderArtWeight, BorderArt, FromTextTop, FromTextBottom, FromTextLeft, FromTextRight, OffsetFrom, InFront, SurroundHeader, SurroundFooter, JoinBorder, LineColor, WhichPages, TL2BRBorder, TR2BLBorder, TL2BRColor, TR2BLColor, TL2BRStyle, TR2BLStyle, TL2BRWeight, TR2BLWeight, ForegroundRGB, BackgroundRGB, TopColorRGB, LeftColorRGB, BottomColorRGB, RightColorRGB, HorizColorRGB, VertColorRGB, TL2BRColorRGB, TR2BLColorRGB, LineColorRGB
- `wdDialogFormatDefineStyleLang` = 186  
  Language, CheckLanguage, Default, NoProof
- `wdDialogFormatPicture` = 187  
  SetSize, CropLeft, CropRight, CropTop, CropBottom, ScaleX, ScaleY, SizeX, SizeY
- `wdDialogToolsLanguage` = 188  
  (none)
- `wdDialogFormatBordersAndShading` = 189  
  ApplyTo, Shadow, TopBorder, LeftBorder, BottomBorder, RightBorder, HorizBorder, VertBorder, TopColor, LeftColor, BottomColor, RightColor, HorizColor, VertColor, FromText, Shading, Foreground, Background, Tab, FineShading, TopStyle, LeftStyle, BottomStyle, RightStyle, HorizStyle, VertStyle, TopWeight, LeftWeight, BottomWeight, RightWeight, HorizWeight, VertWeight, BorderObjectType, BorderArtWeight, BorderArt, FromTextTop, FromTextBottom, FromTextLeft, FromTextRight, OffsetFrom, InFront, SurroundHeader, SurroundFooter, JoinBorder, LineColor, WhichPages, TL2BRBorder, TR2BLBorder, TL2BRColor, TR2BLColor, TL2BRStyle, TR2BLStyle, TL2BRWeight, TR2BLWeight, ForegroundRGB, BackgroundRGB, TopColorRGB, LeftColorRGB, BottomColorRGB, RightColorRGB, HorizColorRGB, VertColorRGB, TL2BRColorRGB, TR2BLColorRGB, LineColorRGB
- `wdDialogFormatFrame` = 190  
  Wrap, WidthRule, FixedWidth, HeightRule, FixedHeight, PositionHorz, PositionHorzRel, DistFromText, PositionVert, PositionVertRel, DistVertFromText, MoveWithText, LockAnchor, RemoveFrame
- `wdDialogToolsThesaurus` = 194  
  (none)
- `wdDialogToolsHyphenation` = 195  
  (none)
- `wdDialogToolsBulletsNumbers` = 196  
  (none)
- `wdDialogToolsHighlightChanges` = 197  
  (none)
- `wdDialogToolsRevisions` = 197  
  (none)
- `wdDialogToolsCompareDocuments` = 198  
  (none)
- `wdDialogTableSort` = 199  
  (none)
- `wdDialogToolsOptionsGeneral` = 203  
  (none)
- `wdDialogToolsOptionsView` = 204  
  (none)
- `wdDialogToolsAdvancedSettings` = 206  
  (none)
- `wdDialogToolsOptionsPrint` = 208  
  (none)
- `wdDialogToolsOptionsSave` = 209  
  (none)
- `wdDialogToolsOptionsSpellingAndGrammar` = 211  
  (none)
- `wdDialogToolsOptionsUserInfo` = 213  
  (none)
- `wdDialogToolsMacroRecord` = 214  
  (none)
- `wdDialogToolsMacro` = 215  
  (none)
- `wdDialogWindowActivate` = 220  
  (none)
- `wdDialogFormatRetAddrFonts` = 221  
  Points, Underline, Color, StrikeThrough, Superscript, Subscript, Hidden, SmallCaps, AllCaps, Spacing, Position, Kerning, KerningMin, Default, Tab, Font, Bold, Italic, DoubleStrikeThrough, Shadow, Outline, Emboss, Engrave, Scale, Animations, CharAccent, FontMajor, FontLowAnsi, FontHighAnsi, CharacterWidthGrid, ColorRGB, UnderlineColor, PointsBi, ColorBi, FontNameBi, BoldBi, ItalicBi, DiacColor
- `wdDialogOrganizer` = 222  
  (none)
- `wdDialogToolsOptionsEdit` = 224  
  (none)
- `wdDialogToolsOptionsFileLocations` = 225  
  (none)
- `wdDialogToolsWordCount` = 228  
  (none)
- `wdDialogControlRun` = 235  
  Application
- `wdDialogInsertPageNumbers` = 294  
  Type, Position, FirstPage
- `wdDialogFormatPageNumber` = 298  
  ChapterNumber, NumRestart, NumFormat, StartingNum, Level, Separator, DoubleQuote, PgNumberingStyle
- `wdDialogCopyFile` = 300  
  FileName, Directory
- `wdDialogFormatChangeCase` = 322  
  Type
- `wdDialogUpdateTOC` = 331  
  (none)
- `wdDialogInsertDatabase` = 341  
  Format, Style, LinkToSource, Connection, SQLStatement, SQLStatement1, PasswordDoc, PasswordDot, DataSource, From, To, IncludeFields, WritePasswordDoc, WritePasswordDot
- `wdDialogTableFormula` = 348  
  (none)
- `wdDialogFormFieldOptions` = 353  
  Entry, Exit, Name, Enable, TextType, TextWidth, TextDefault, TextFormat, CheckSize, CheckWidth, CheckDefault, Type, OwnHelp, HelpText, OwnStat, StatText, Calculate
- `wdDialogInsertCaption` = 357  
  Label, TitleAutoText, Title, Delete, Position, AutoCaption, ExcludeLabel
- `wdDialogInsertCaptionNumbering` = 358  
  Label, FormatNumber, ChapterNumber, Level, Separator, CapNumberingStyle
- `wdDialogInsertAutoCaption` = 359  
  Clear, ClearAll, Object, Label, Position
- `wdDialogFormFieldHelp` = 361  
  (none)
- `wdDialogInsertCrossReference` = 367  
  ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperLink, InsertPosition, SeparateNumbers, SeparatorCharacters
- `wdDialogInsertFootnote` = 370  
  Reference, NoteType, Symbol, FootNumberAs, EndNumberAs, FootnotesAt, EndnotesAt, FootNumberingStyle, EndNumberingStyle, FootStartingNum, FootRestartNum, EndStartingNum, EndRestartNum, ApplyPropsTo
- `wdDialogNoteOptions` = 373  
  (none)
- `wdDialogToolsAutoCorrect` = 378  
  (none)
- `wdDialogToolsOptionsTrackChanges` = 386  
  (none)
- `wdDialogConvertObject` = 392  
  IconNumber, ActivateAs, IconFileName, Caption, Class, DisplayIcon, Floating
- `wdDialogInsertAddCaption` = 402  
  Name
- `wdDialogConnect` = 420  
  Drive, Path, Password
- `wdDialogToolsCustomizeKeyboard` = 432  
  (none)
- `wdDialogToolsCustomizeMenus` = 433  
  (none)
- `wdDialogToolsMergeDocuments` = 435  
  (none)
- `wdDialogMarkTableOfContentsEntry` = 442  
  (none)
- `wdDialogFileMacPageSetupGX` = 444  
  Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `wdDialogFilePrintOneCopy` = 445  
  Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `wdDialogEditFrame` = 458  
  Wrap, WidthRule, FixedWidth, HeightRule, FixedHeight, PositionHorz, PositionHorzRel, DistFromText, PositionVert, PositionVertRel, DistVertFromText, MoveWithText, LockAnchor, RemoveFrame
- `wdDialogMarkCitation` = 463  
  (none)
- `wdDialogTableOfContentsOptions` = 470  
  (none)
- `wdDialogInsertTableOfAuthorities` = 471  
  Outline, Fields, From, To, TableId, AddedStyles, Caption, HeadingSeparator, Replace, MarkEntry, AutoMark, MarkCitation, Type, RightAlignPageNumbers, Passim, KeepFormatting, Columns, Category, Label, ShowPageNumbers, AccentedLetters, Filter, SortBy, Leader, TOCUseHyperlinks, TOCHidePageNumInWeb, IndexLanguage, UseOutlineLevel
- `wdDialogInsertTableOfFigures` = 472  
  Outline, Fields, From, To, TableId, AddedStyles, Caption, HeadingSeparator, Replace, MarkEntry, AutoMark, MarkCitation, Type, RightAlignPageNumbers, Passim, KeepFormatting, Columns, Category, Label, ShowPageNumbers, AccentedLetters, Filter, SortBy, Leader, TOCUseHyperlinks, TOCHidePageNumInWeb, IndexLanguage, UseOutlineLevel
- `wdDialogInsertIndexAndTables` = 473  
  Outline, Fields, From, To, TableId, AddedStyles, Caption, HeadingSeparator, Replace, MarkEntry, AutoMark, MarkCitation, Type, RightAlignPageNumbers, Passim, KeepFormatting, Columns, Category, Label, ShowPageNumbers, AccentedLetters, Filter, SortBy, Leader, TOCUseHyperlinks, TOCHidePageNumInWeb, IndexLanguage, UseOutlineLevel
- `wdDialogInsertFormField` = 483  
  Entry, Exit, Name, Enable, TextType, TextWidth, TextDefault, TextFormat, CheckSize, CheckWidth, CheckDefault, Type, OwnHelp, HelpText, OwnStat, StatText, Calculate
- `wdDialogFormatDropCap` = 488  
  Position, Font, DropHeight, DistFromText
- `wdDialogToolsCreateLabels` = 489  
  (none)
- `wdDialogToolsProtectDocument` = 503  
  (none)
- `wdDialogFormatStyleGallery` = 505  
  Template, Preview
- `wdDialogToolsAcceptRejectChanges` = 506  
  (none)
- `wdDialogHelpWordPerfectHelpOptions` = 511  
  CommandKeyHelp, DocNavKeys, MouseSimulation, DemoGuidance, DemoSpeed, HelpType
- `wdDialogToolsUnprotectDocument` = 521  
  (none)
- `wdDialogToolsOptionsCompatibility` = 525  
  (none)
- `wdDialogTableOfCaptionsOptions` = 551  
  (none)
- `wdDialogTableAutoFormat` = 563  
  (none)
- `wdDialogMailMergeFindRecord` = 569  
  (none)
- `wdDialogReviewAfmtRevisions` = 570  
  (none)
- `wdDialogViewZoom` = 577  
  (none)
- `wdDialogToolsProtectSection` = 578  
  (none)
- `wdDialogFontSubstitution` = 581  
  UnavailableFont, SubstituteFont
- `wdDialogInsertSubdocument` = 583  
  Name, ConfirmConversions, ReadOnly, LinkToSource, AddToMru, PasswordDoc, PasswordDot, Revert, WritePasswordDoc, WritePasswordDot, Connection, SQLStatement, SQLStatement1, Format, Encoding, Visible, OpenExclusive, OpenAndRepair, SubType, DocumentDirection, NoEncodingDialog, XMLTransform
- `wdDialogNewToolbar` = 586  
  (none)
- `wdDialogToolsEnvelopesAndLabels` = 607  
  (none)
- `wdDialogFormatCallout` = 610  
  Type, Gap, Angle, Drop, Length, Border, AutoAttach, Accent
- `wdDialogTableFormatCell` = 612  
  (none)
- `wdDialogToolsCustomizeMenuBar` = 615  
  (none)
- `wdDialogFileRoutingSlip` = 624  
  Subject, Message, AllAtOnce, ReturnWhenDone, TrackStatus, Protect, AddSlip, RouteDocument, AddRecipient, OldRecipient, ResetSlip, ClearSlip, ClearRecipients, Address
- `wdDialogEditTOACategory` = 625  
  Category, CategoryName
- `wdDialogToolsManageFields` = 631  
  (none)
- `wdDialogDrawSnapToGrid` = 633  
  SnapToGrid, XGrid, YGrid, XOrigin, YOrigin, SnapToShapes, XGridDisplay, YGridDisplay, FollowMargins, ViewGridLines, DefineLineBasedOnGrid
- `wdDialogDrawAlign` = 634  
  Horizontal, Vertical, RelativeTo
- `wdDialogMailMergeCreateDataSource` = 642  
  FileName, PasswordDoc, PasswordDot, HeaderRecord, MSQuery, SQLStatement, SQLStatement1, Connection, LinkToSource, WritePasswordDoc
- `wdDialogMailMergeCreateHeaderSource` = 643  
  FileName, PasswordDoc, PasswordDot, HeaderRecord, MSQuery, SQLStatement, SQLStatement1, Connection, LinkToSource, WritePasswordDoc
- `wdDialogMailMerge` = 676  
  CheckErrors, Destination, MergeRecords, From, To, Suppression, MailMerge, QueryOptions, MailSubject, MailAsAttachment, MailAddress
- `wdDialogMailMergeCheck` = 677  
  CheckErrors
- `wdDialogMailMergeHelper` = 680  
  (none)
- `wdDialogMailMergeQueryOptions` = 681  
  (none)
- `wdDialogFileMacPageSetup` = 685  
  Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `wdDialogListCommands` = 723  
  ListType
- `wdDialogEditCreatePublisher` = 732  
  Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `wdDialogEditSubscribeTo` = 733  
  Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `wdDialogEditPublishOptions` = 735  
  Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `wdDialogEditSubscribeOptions` = 736  
  Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `wdDialogFileMacCustomPageSetupGX` = 737  
  Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `wdDialogToolsOptionsTypography` = 739  
  (none)
- `wdDialogToolsAutoCorrectExceptions` = 762  
  (none)
- `wdDialogToolsOptionsAutoFormatAsYouType` = 778  
  (none)
- `wdDialogMailMergeUseAddressBook` = 779  
  (none)
- `wdDialogToolsHangulHanjaConversion` = 784  
  (none)
- `wdDialogToolsOptionsFuzzy` = 790  
  (none)
- `wdDialogEditGoToOld` = 811  
  (none)
- `wdDialogInsertNumber` = 812  
  NumPic
- `wdDialogLetterWizard` = 821  
  SenderCity, DateFormat, IncludeHeaderFooter, LetterStyle, Letterhead, LetterheadLocation, LetterheadSize, RecipientName, RecipientAddress, Salutation, SalutationType, RecipientGender, RecipientReference, MailingInstructions, AttentionLine, LetterSubject, CCList, SenderName, ReturnAddress, Closing, SenderJobTitle, SenderCompany, SenderInitials, EnclosureNumber, PageDesign, InfoBlock, SenderGender, ReturnAddressSF, RecipientCode, SenderCode, SenderReference
- `wdDialogFormatBulletsAndNumbering` = 824  
  (none)
- `wdDialogToolsSpellingAndGrammar` = 828  
  (none)
- `wdDialogToolsCreateDirectory` = 833  
  (none)
- `wdDialogTableWrapping` = 854  
  (none)
- `wdDialogFormatTheme` = 855  
  (none)
- `wdDialogTableProperties` = 861  
  (none)
- `wdDialogEmailOptions` = 863  
  (none)
- `wdDialogCreateAutoText` = 872  
  (none)
- `wdDialogToolsAutoSummarize` = 874  
  (none)
- `wdDialogToolsGrammarSettings` = 885  
  (none)
- `wdDialogEditGoTo` = 896  
  Find, Replace, Direction, MatchCase, WholeWord, PatternMatch, SoundsLike, FindNext, ReplaceOne, ReplaceAll, Format, Wrap, FindAllWordForms, MatchByte, FuzzyFind, Destination, CorrectEnd, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl
- `wdDialogWebOptions` = 898  
  (none)
- `wdDialogInsertHyperlink` = 925  
  (none)
- `wdDialogToolsAutoManager` = 915  
  (none)
- `wdDialogFileVersions` = 945  
  AutoVersion, VersionDesc
- `wdDialogToolsOptionsAutoFormat` = 959  
  (none)
- `wdDialogFormatDrawingObject` = 960  
  Left, PositionHorzRel, Top, PositionVertRel, LockAnchor, FloatOverText, LayoutInCell, WrapSide, TopDistanceFromText, BottomDistanceFromText, LeftDistanceFromText, RightDistanceFromText, Wrap, WordWrap, AutoSize, HRWidthType, HRHeight, HRNoshade, HRAlign, Text, AllowOverlap, HorizRule
- `wdDialogToolsOptions` = 974  
  (none)
- `wdDialogFitText` = 983  
  FitTextWidth
- `wdDialogEditAutoText` = 985  
  Name, Context, InsertAs, Insert, Add, Define, InsertAsText, Delete, CompleteAT
- `wdDialogPhoneticGuide` = 986  
  (none)
- `wdDialogToolsDictionary` = 989  
  (none)
- `wdDialogFileSaveVersion` = 1007  
  (none)
- `wdDialogToolsOptionsBidi` = 1029  
  (none)
- `wdDialogFrameSetProperties` = 1074  
  (none)
- `wdDialogTableTableOptions` = 1080  
  (none)
- `wdDialogTableCellOptions` = 1081  
  (none)
- `wdDialogIMESetDefault` = 1094  
  (none)
- `wdDialogTCSCTranslator` = 1156  
  (none)
- `wdDialogHorizontalInVertical` = 1160  
  (none)
- `wdDialogTwoLinesInOne` = 1161  
  (none)
- `wdDialogFormatEncloseCharacters` = 1162  
  Style, Text, Enclosure
- `wdDialogConsistencyChecker` = 1121  
  (none)
- `wdDialogToolsOptionsSmartTag` = 1395  
  (none)
- `wdDialogFormatStylesCustom` = 1248  
  (none)
- `wdDialogCSSLinks` = 1261  
  LinkStyles
- `wdDialogInsertWebComponent` = 1324  
  IconNumber, FileName, Link, DisplayIcon, Tab, Class, IconFileName, Caption, Floating
- `wdDialogToolsOptionsEditCopyPaste` = 1356  
  (none)
- `wdDialogToolsOptionsSecurity` = 1361  
  (none)
- `wdDialogSearch` = 1363  
  (none)
- `wdDialogShowRepairs` = 1381  
  (none)
- `wdDialogMailMergeInsertAsk` = 4047  
  (none)
- `wdDialogMailMergeInsertFillIn` = 4048  
  (none)
- `wdDialogMailMergeInsertIf` = 4049  
  (none)
- `wdDialogMailMergeInsertNextIf` = 4053  
  (none)
- `wdDialogMailMergeInsertSet` = 4054  
  (none)
- `wdDialogMailMergeInsertSkipIf` = 4055  
  (none)
- `wdDialogMailMergeFieldMapping` = 1304  
  (none)
- `wdDialogMailMergeInsertAddressBlock` = 1305  
  (none)
- `wdDialogMailMergeInsertGreetingLine` = 1306  
  (none)
- `wdDialogMailMergeInsertFields` = 1307  
  (none)
- `wdDialogMailMergeRecipients` = 1308  
  (none)
- `wdDialogMailMergeFindRecipient` = 1326  
  (none)
- `wdDialogMailMergeSetDocumentType` = 1339  
  (none)
- `wdDialogLabelOptions` = 1367  
  (none)
- `wdDialogXMLElementAttributes` = 1460  
  (none)
- `wdDialogSchemaLibrary` = 1417  
  (none)
- `wdDialogPermission` = 1469  
  (none)
- `wdDialogMyPermission` = 1437  
  (none)
- `wdDialogXMLOptions` = 1425  
  (none)
- `wdDialogFormattingRestrictions` = 1427  
  (none)
- `wdDialogSourceManager` = 1920  
  (none)
- `wdDialogCreateSource` = 1922  
  (none)
- `wdDialogDocumentInspector` = 1482  
  (none)
- `wdDialogStyleManagement` = 1948  
  (none)
- `wdDialogInsertSource` = 2120  
  (none)
- `wdDialogOMathRecognizedFunctions` = 2165  
  (none)
- `wdDialogInsertPlaceholder` = 2348  
  (none)
- `wdDialogBuildingBlockOrganizer` = 2067  
  (none)
- `wdDialogContentControlProperties` = 2394  
  (none)
- `wdDialogCompatibilityChecker` = 2439
- `wdDialogExportAsFixedFormat` = 2349
- `wdDialogFileNew2007` = 1116
