# Options

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209B7-0000-0000-C000-000000000046}  

Represents application and document options in Word. Many of the properties for the Options object correspond to items in the Options dialog box.

**Remarks:** Use the Options property to return the Options object. The following example sets three application options for Word.

## Properties (248)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Options object.
- `AllowAccentedUppercase As Boolean  (read/write)`  
  True if accents are retained when a French language character is changed to uppercase. Read/write Boolean.
- `Pagination As Boolean  (read/write)`  
  True if Microsoft Word repaginates documents in the background. Read/write Boolean.
- `EnableSound As Boolean  (read/write)`  
  True if Word makes the computer respond with a sound whenever an error occurs. Read/write Boolean.
- `ConfirmConversions As Boolean  (read/write)`  
  True if Word displays the Convert File dialog box before it opens or inserts a file that isn't a Word document or template. In the Convert File dialog box, the user chooses the format to convert the file from. Read/write Boolean.
- `UpdateLinksAtOpen As Boolean  (read/write)`  
  True if Microsoft Word automatically updates all embedded OLE links in a document when it is opened. Read/write Boolean.
- `SendMailAttach As Boolean  (read/write)`  
  True if the Send To command on the File menu inserts the active document as an attachment to a mail message. Read/write Boolean.
- `MeasurementUnit As WdMeasurementUnits  (read/write)`  
  Returns or sets the standard measurement unit for Microsoft Word. Read/write WdMeasurementUnits.
- `ButtonFieldClicks As Long  (read/write)`  
  Returns or sets the number of clicks (either one or two) required to run a GOTOBUTTON or MACROBUTTON field. Read/write Long.
- `ShortMenuNames As Boolean  (read/write)`
- `RTFInClipboard As Boolean  (read/write)`  
  You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about the RTFInClipboard property for the Options object, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `UpdateFieldsAtPrint As Boolean  (read/write)`  
  True if Microsoft Word updates fields automatically before printing a document. Read/write Boolean.
- `PrintProperties As Boolean  (read/write)`  
  True if Microsoft Word prints document summary information on a separate page at the end of the document. Read/write Boolean.
- `PrintFieldCodes As Boolean  (read/write)`  
  True if Microsoft Word prints field codes instead of field results. Read/write Boolean.
- `PrintComments As Boolean  (read/write)`  
  True if Microsoft Word prints comments, starting on a new page at the end of the document. Read/write Boolean.
- `PrintHiddenText As Boolean  (read/write)`  
  True if hidden text is printed. Read/write Boolean.
- `EnvelopeFeederInstalled As Boolean  (read-only)`  
  True if the current printer has a special feeder for envelopes. Read-only Boolean.
- `UpdateLinksAtPrint As Boolean  (read/write)`  
  True if Microsoft Word updates embedded links to other files before printing a document. Read/write Boolean.
- `PrintBackground As Boolean  (read/write)`  
  True if Microsoft Word prints in the background. Read/write Boolean.
- `PrintDrawingObjects As Boolean  (read/write)`  
  True if Microsoft Word prints drawing objects. Read/write Boolean.
- `DefaultTray As String  (read/write)`  
  Returns or sets the default tray your printer uses to print documents. Read/write String.
- `DefaultTrayID As Long  (read/write)`  
  Returns or sets the default tray your printer uses to print documents. Read/write WdPaperTray.
- `CreateBackup As Boolean  (read/write)`  
  True if Word creates a backup copy each time a document is saved. Read/write Boolean.
- `SavePropertiesPrompt As Boolean  (read/write)`  
  True if Microsoft Word prompts for document property information when saving a new document. Read/write Boolean.
- `SaveNormalPrompt As Boolean  (read/write)`  
  True if Microsoft Word prompts the user for confirmation to save changes to the Normal template before it closes. Read/write Boolean.
- `SaveInterval As Long  (read/write)`  
  Returns or sets the time interval in minutes for saving AutoRecover information. Read/write Long.
- `BackgroundSave As Boolean  (read/write)`  
  True if Word saves documents in the background. When Word is saving in the background, users can continue to type and to choose commands. Read/write Boolean.
- `InsertedTextMark As WdInsertedTextMark  (read/write)`  
  Returns or sets how Microsoft Word formats inserted text while change tracking is enabled (the TrackRevisions property is True). Read/write WdInsertedTextMark.
- `DeletedTextMark As WdDeletedTextMark  (read/write)`  
  Returns or sets the format of text that is deleted while change tracking is enabled. Read/write WdDeletedTextMark.
- `RevisedLinesMark As WdRevisedLinesMark  (read/write)`  
  Returns or sets the placement of changed lines in a document with tracked changes. Read/write WdRevisedLinesMark.
- `InsertedTextColor As WdColorIndex  (read/write)`  
  Returns or sets the color of text that is inserted while change tracking is enabled. Read/write WdColorIndex.
- `DeletedTextColor As WdColorIndex  (read/write)`  
  Returns or sets the color of text that is deleted while change tracking is enabled. Read/write WdColorIndex.
- `RevisedLinesColor As WdColorIndex  (read/write)`  
  Returns or sets the color of changed lines in a document with tracked changes. Read/write WdColorIndex.
- `DefaultFilePath As String  (read/write)`  
  Returns or sets default folders for items such as documents, templates, and graphics. Read/write String.
- `Overtype As Boolean  (read/write)`  
  True if Overtype mode is active. Read/write Boolean.
- `ReplaceSelection As Boolean  (read/write)`  
  True if the result of typing or pasting replaces the selection. Read/write Boolean.
- `AllowDragAndDrop As Boolean  (read/write)`  
  True if dragging can be used to move or copy a selection. Read/write Boolean.
- `AutoWordSelection As Boolean  (read/write)`  
  True if dragging selects one word at a time instead of one character at a time. Read/write Boolean.
- `INSKeyForPaste As Boolean  (read/write)`  
  True if the INS key can be used for pasting the Clipboard contents. Read/write Boolean.
- `SmartCutPaste As Boolean  (read/write)`  
  True if Microsoft Word automatically adjusts the spacing between words and punctuation when cutting and pasting occurs. Read/write Boolean.
- `TabIndentKey As Boolean  (read/write)`  
  True if the TAB and BACKSPACE keys can be used to increase and decrease, respectively, the left indent of paragraphs and if the BACKSPACE key can be used to change right-aligned paragraphs to centered paragraphs and centered paragraphs to left-aligned paragraphs. Read/write Boolean.
- `PictureEditor As String  (read/write)`  
  Returns or sets the name of the application to use to edit pictures. Read/write String.
- `AnimateScreenMovements As Boolean  (read/write)`  
  True if Word animates mouse movements, uses animated cursors, and animates actions such as background saving and find and replace operations. Read/write Boolean.
- `RevisedPropertiesMark As WdRevisedPropertiesMark  (read/write)`  
  Returns or sets the mark used to show formatting changes while change tracking is enabled. Read/write WdRevisedPropertiesMark.
- `RevisedPropertiesColor As WdColorIndex  (read/write)`  
  Returns or sets the color used to mark formatting changes while change tracking is enabled. Read/write WdColorIndex.
- `SnapToGrid As Boolean  (read/write)`  
  True if AutoShapes or East Asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized. Read/write Boolean.
- `SnapToShapes As Boolean  (read/write)`  
  True if Word automatically aligns AutoShapes or East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other AutoShapes or East Asian characters. Read/write Boolean.
- `GridDistanceHorizontal As Single  (read/write)`  
  Returns or sets the amount of horizontal space between the invisible gridlines that Word uses when you draw, move, and resize AutoShapes or East Asian characters in new documents. Read/write Single.
- `GridDistanceVertical As Single  (read/write)`  
  Returns or sets the amount of vertical space between the invisible gridlines that Word uses when you draw, move, and resize AutoShapes or East Asian characters in new documents. Read/write Single.
- `GridOriginHorizontal As Single  (read/write)`  
  Returns or sets the point, relative to the left edge of the page, where you want the invisible grid for drawing, moving, and resizing AutoShapes or East Asian characters to begin in new documents. Read/write Single.
- `GridOriginVertical As Single  (read/write)`  
  Returns or sets the point, relative to the top of the page, where you want the invisible grid for drawing, moving, and resizing AutoShapes or East Asian characters to begin in new documents. Read/write Single.
- `InlineConversion As Boolean  (read/write)`  
  True if Microsoft Word displays an unconfirmed character string in the Japanese Input Method Editor (IME) as an insertion between existing (confirmed) character strings. Read/write Boolean.
- `IMEAutomaticControl As Boolean  (read/write)`  
  True if Microsoft Word is set to automatically open and close the Japanese Input Method Editor (IME). Read/write Boolean.
- `AutoFormatApplyHeadings As Boolean  (read/write)`  
  True if styles are automatically applied to headings when Word formats a document or range automatically. Read/write Boolean.
- `AutoFormatApplyLists As Boolean  (read/write)`  
  True if styles are automatically applied to lists when Word formats a document or range automatically. Read/write Boolean.
- `AutoFormatApplyBulletedLists As Boolean  (read/write)`  
  True if characters (such as asterisks, hyphens, and greater-than signs) at the beginning of list paragraphs are replaced with bullets from the Bullets and Numbering dialog box (Format menu) when Word formats a document or range automatically. Read/write Boolean.
- `AutoFormatApplyOtherParas As Boolean  (read/write)`  
  True if styles are automatically applied to paragraphs that aren't headings or list items when Word formats a document or range automatically. Read/write Boolean.
- `AutoFormatReplaceQuotes As Boolean  (read/write)`  
  True if straight quotation marks are automatically changed to smart (curly) quotation marks when Word formats a document or range automatically. Read/write Boolean.
- `AutoFormatReplaceSymbols As Boolean  (read/write)`  
  True if two consecutive hyphens (--) are replaced by an en dash (-) or an em dash (-) when Word formats a document or range automatically. Read/write Boolean.
- `AutoFormatReplaceOrdinals As Boolean  (read/write)`  
  True if the ordinal number suffixes "st", "nd", "rd", and "th" are replaced with the same letters in superscript when Word formats a document or range automatically. For example, "1st" is replaced with "1" followed by "st" formatted as superscript. Read/write Boolean.
- `AutoFormatReplaceFractions As Boolean  (read/write)`  
  True if typed fractions are replaced with fractions from the current character set when Word formats a document or range automatically. For example, "1/2" is replaced with "." Read/write Boolean.
- `AutoFormatReplacePlainTextEmphasis As Boolean  (read/write)`  
  True if manual emphasis characters are replaced with character formatting when Word formats a document or range automatically. For example, "bold" is changed to "bold" and "underline" is changed to "underline." Read/write Boolean.
- `AutoFormatPreserveStyles As Boolean  (read/write)`  
  True if previously applied styles are preserved when Word formats a document or range automatically. Read/write Boolean.
- `AutoFormatAsYouTypeApplyHeadings As Boolean  (read/write)`  
  True if styles are automatically applied to headings as you type. Read/write Boolean.
- `AutoFormatAsYouTypeApplyBorders As Boolean  (read/write)`  
  True if a series of three or more hyphens (-), equal signs (=), or underscore characters (_) are automatically replaced by a specific border line when the ENTER key is pressed. Read/write Boolean.
- `AutoFormatAsYouTypeApplyBulletedLists As Boolean  (read/write)`  
  True if bullet characters (such as asterisks, hyphens, and greater-than signs) are replaced with bullets from the Bullets And Numbering dialog box (Format menu) as you type. Read/write Boolean.
- `AutoFormatAsYouTypeApplyNumberedLists As Boolean  (read/write)`  
  True if paragraphs are automatically formatted as numbered lists with a numbering scheme from the Bullets and Numbering dialog box (Format menu), according to what's typed. For example, if a paragraph starts with "1.1" and a tab character, Word automatically inserts "1.2" and a tab character after the ENTER key is pressed. Read/write Boolean.
- `AutoFormatAsYouTypeReplaceQuotes As Boolean  (read/write)`  
  True if straight quotation marks are automatically changed to smart (curly) quotation marks as you type. Read/write Boolean.
- `AutoFormatAsYouTypeReplaceSymbols As Boolean  (read/write)`  
  True if two consecutive hyphens (--) are replaced with an en dash (-) or an em dash (-) as you type. Read/write Boolean.If the hyphens are typed with leading and trailing spaces, Word replaces the hyphens with an en dash; if there are no trailing spaces, the hyphens are replaced with an em dash.
- `AutoFormatAsYouTypeReplaceOrdinals As Boolean  (read/write)`  
  True if the ordinal number suffixes "st", "nd", "rd", and "th" are replaced with the same letters in superscript as you type. For example, "1st" is replaced with "1" followed by "st" formatted as superscript. Read/write Boolean.
- `AutoFormatAsYouTypeReplaceFractions As Boolean  (read/write)`  
  True if typed fractions are replaced with fractions from the current character set as you type. For example, "1/2" is replaced with "." Read/write Boolean.
- `AutoFormatAsYouTypeReplacePlainTextEmphasis As Boolean  (read/write)`  
  True if manual emphasis characters are automatically replaced with character formatting as you type. For example, "bold" is changed to " bold " and "underline" is changed to "underline." Read/write Boolean.
- `AutoFormatAsYouTypeFormatListItemBeginning As Boolean  (read/write)`  
  True if Word repeats character formatting applied to the beginning of a list item to the next list item. Read/write Boolean.
- `AutoFormatAsYouTypeDefineStyles As Boolean  (read/write)`  
  True if Word automatically creates new styles based on manual formatting. Read/write Boolean.
- `AutoFormatPlainTextWordMail As Boolean  (read/write)`  
  True if Word automatically formats plain-text email messages when you open them in Word. Read/write Boolean.
- `AutoFormatAsYouTypeReplaceHyperlinks As Boolean  (read/write)`  
  True if email addresses, server and share names (also known as UNC paths), and Internet addresses (also known as URLs) are automatically changed to hyperlinks as you type. Read/write Boolean.
- `AutoFormatReplaceHyperlinks As Boolean  (read/write)`  
  True if email addresses, server and share names (also known as UNC paths), and Internet addresses (also known as URLs) are automatically formatted whenever Word AutoFormats a document or range. Read/write Boolean.
- `DefaultHighlightColorIndex As WdColorIndex  (read/write)`  
  Returns or sets the color used to highlight text formatted with the Highlight button (Formatting toolbar). Read/write WdColorIndex.
- `DefaultBorderLineStyle As WdLineStyle  (read/write)`  
  Returns or sets the default border line style. Read/write WdLineStyle.
- `CheckSpellingAsYouType As Boolean  (read/write)`  
  True if Microsoft Word checks spelling and marks errors automatically as you type. Read/write Boolean.
- `CheckGrammarAsYouType As Boolean  (read/write)`  
  True if Word checks grammar and marks errors automatically as you type. Read/write Boolean.
- `IgnoreInternetAndFileAddresses As Boolean  (read/write)`  
  True if file name extensions, MS-DOS paths, email addresses, server and share names (also known as UNC paths), and Internet addresses (also known as URLs) are ignored while checking spelling. Read/write Boolean.
- `ShowReadabilityStatistics As Boolean  (read/write)`  
  True if Microsoft Word displays a list of summary statistics, including measures of readability, when it has finished checking grammar. Read/write Boolean.
- `IgnoreUppercase As Boolean  (read/write)`  
  True if words in all uppercase letters are ignored while checking spelling. Read/write Boolean.
- `IgnoreMixedDigits As Boolean  (read/write)`  
  True if words that contain numbers are ignored while checking spelling. Read/write Boolean.
- `SuggestFromMainDictionaryOnly As Boolean  (read/write)`  
  True if Microsoft Word draws spelling suggestions from the main dictionary only. Read/write Boolean.
- `SuggestSpellingCorrections As Boolean  (read/write)`  
  True if Microsoft Word always suggests alternative spellings for each misspelled word when checking spelling. Read/write Boolean.
- `DefaultBorderLineWidth As WdLineWidth  (read/write)`  
  Returns or sets the default line width of borders. Read/write WdLineWidth.
- `CheckGrammarWithSpelling As Boolean  (read/write)`  
  True if Word checks grammar while checking spelling. Read/write Boolean.
- `DefaultOpenFormat As WdOpenFormat  (read/write)`  
  Returns or sets the default file converter used to open documents. Can be a number returned by the OpenFormat property, or one of the following WdOpenFormat constants.
- `PrintDraft As Boolean  (read/write)`  
  True if Microsoft Word prints using minimal formatting. Read/write Boolean.
- `PrintReverse As Boolean  (read/write)`  
  True if Microsoft Word prints pages in reverse order. Read/write Boolean.
- `MapPaperSize As Boolean  (read/write)`  
  True if documents formatted for another country's/region's standard paper size (for example, A4) are automatically adjusted so that they're printed correctly on your country's/region's standard paper size (for example, Letter). Read/write Boolean.
- `AutoFormatAsYouTypeApplyTables As Boolean  (read/write)`  
  True if Word automatically creates a table when you type a plus sign, a series of hyphens, another plus sign, and so on, and then press ENTER. The plus signs become the column borders, and the hyphens become the column widths. Read/write Boolean.
- `AutoFormatApplyFirstIndents As Boolean  (read/write)`  
  True if Microsoft Word replaces a space entered at the beginning of a paragraph with a first-line indent when Word formats a document or range automatically. Read/write Boolean.
- `AutoFormatMatchParentheses As Boolean  (read/write)`  
  True if improperly paired parentheses are corrected when Microsoft Word formats a document or range automatically. Read/write Boolean.
- `AutoFormatReplaceFarEastDashes As Boolean  (read/write)`  
  True if long vowel sound and dash use is corrected when Microsoft Word formats a document or range automatically. Read/write Boolean.
- `AutoFormatDeleteAutoSpaces As Boolean  (read/write)`  
  True if spaces inserted between Japanese and Latin text will be deleted when Microsoft Word formats a document or range automatically. Read/write Boolean.
- `AutoFormatAsYouTypeApplyFirstIndents As Boolean  (read/write)`  
  True for Microsoft Word to automatically replace a space entered at the beginning of a paragraph with a first-line indent. Read/write.
- `AutoFormatAsYouTypeApplyDates As Boolean  (read/write)`  
  True for Microsoft Word to automatically apply the Date style to dates as you type. Read/write.
- `AutoFormatAsYouTypeApplyClosings As Boolean  (read/write)`  
  True for Microsoft Word to automatically apply the Closing style to letter closings as you type. Read/write Boolean.
- `AutoFormatAsYouTypeMatchParentheses As Boolean  (read/write)`  
  True for Microsoft Word to automatically correct improperly paired parentheses. Read/write.
- `AutoFormatAsYouTypeReplaceFarEastDashes As Boolean  (read/write)`  
  True for Microsoft Word to automatically correct long vowel sounds and dashes. Read/write.
- `AutoFormatAsYouTypeDeleteAutoSpaces As Boolean  (read/write)`  
  True for Microsoft Word to automatically delete spaces inserted between Japanese and Latin text as you type. Read/write.
- `AutoFormatAsYouTypeInsertClosings As Boolean  (read/write)`  
  True for Microsoft Word to automatically insert the corresponding memo closing when the user enters a memo heading. Read/write.
- `AutoFormatAsYouTypeAutoLetterWizard As Boolean  (read/write)`  
  True for Microsoft Word to automatically start the Letter Wizard when the user enters a letter salutation or closing. Read/write.
- `AutoFormatAsYouTypeInsertOvers As Boolean  (read/write)`  
  True for Microsoft Word to automatically insert "" when the user enters "" or "". Read/write Boolean.
- `DisplayGridLines As Boolean  (read/write)`  
  True if Microsoft Word displays the document grid. This property is the equivalent of the Gridlines command on the View menu. Read/write Boolean.
- `MatchFuzzyCase As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between uppercase and lowercase letters during a search. Read/write Boolean.
- `MatchFuzzyByte As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between full-width and half-width characters (Latin or Japanese) during a search. Read/write Boolean.
- `MatchFuzzyHiragana As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between hiragana and katakana during a search. Read/write Boolean.
- `MatchFuzzySmallKana As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between diphthongs and double consonants during a search. Read/write Boolean.
- `MatchFuzzyDash As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between minus signs, long vowel sounds, and dashes during a search. Read/write Boolean.
- `MatchFuzzyIterationMark As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between types of repetition marks during a search. Read/write Boolean.
- `MatchFuzzyKanji As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between standard and nonstandard kanji ideography during a search. Read/write Boolean.
- `MatchFuzzyOldKana As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between new kana and old kana characters during a search. Read/write Boolean.
- `MatchFuzzyProlongedSoundMark As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between short and long vowel sounds during a search. Read/write Boolean.
- `MatchFuzzyDZ As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between " !Screenshot of the first symbol in the example." and " !Screenshot of the second symbol in the example." and between " !Screenshot of the third symbol in the example." and " !Screenshot of the fourth symbol in the example." during a search. Read/write Boolean.
- `MatchFuzzyBV As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between " !Screenshot of symbol #1 in the example." and " !Screenshot of symbol #2 in the example. !Screenshot of symbol #3 in the example." and between " !Screenshot of symbol #4 in the example." and " !Screenshot of symbol #5 in the example. !Screenshot of symbol #6 in the example." during a search. Read/write Boolean.
- `MatchFuzzyTC As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between " !Screenshot of the first symbol in the example. !Screenshot of the second symbol in the example.", " !Screenshot of the third symbol in the example. !Screenshot of the fourth symbol in the example.", and " !Screenshot of the fifth symbol in the example.", and between " !Screenshot of the sixth symbol in the example. !Screenshot of the seventh symbol in the example." and " !Screenshot of the eighth symbol in the example." during a search. Read/write Boolean.
- `MatchFuzzyHF As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between " !Screenshot of the first symbol in the example. !Screenshot of the second symbol in the example." and " !Screenshot of the third symbol in the example. !Screenshot of the fourth symbol in the example." and between " !Screenshot of the fifth symbol in the example. !Screenshot of the sixth symbol in the example." and " !Screenshot of the seventh symbol in the example. !Screenshot of the eighth symbol in the example." during a search. Read/write Boolean.
- `MatchFuzzyZJ As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between " !Screenshot of symbol #1 in the example." and " !Screenshot of symbol #2 in the example. !Screenshot of symbol #3 in the example." and between " !Screenshot of symbol #4 in the example." and " !Screenshot of symbol #5 in the example. !Screenshot of symbol #6 in the example." during a search. Read/write Boolean.
- `MatchFuzzyAY As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between " !Screenshot of the first symbol in the example." and " !Screenshot of the second symbol in the example." following !Screenshot of the third symbol in the example.-row and !Screenshot of the fourth symbol in the example.-row characters during a search. Read/write Boolean.
- `MatchFuzzyKiKu As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between " !Screenshot of symbol #1 in the example." and " !Screenshot of symbol #2 in the example." before !Screenshot of symbol #3 in the example.-row characters during a search. Read/write Boolean.
- `MatchFuzzyPunctuation As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between types of punctuation marks during a search. Read/write Boolean.
- `MatchFuzzySpace As Boolean  (read/write)`  
  True if Microsoft Word ignores the distinction between space markers used during a search. Read/write Boolean.
- `ApplyFarEastFontsToAscii As Boolean  (read/write)`  
  True if Microsoft Word applies East Asian fonts to Latin text. Read/write Boolean.
- `ConvertHighAnsiToFarEast As Boolean  (read/write)`  
  True if Microsoft Word converts text that is associated with an East Asian font to the appropriate font when it opens a document. Read/write Boolean.
- `PrintOddPagesInAscendingOrder As Boolean  (read/write)`  
  True if Microsoft Word prints odd pages in ascending order during manual duplex printing. Read/write Boolean.
- `PrintEvenPagesInAscendingOrder As Boolean  (read/write)`  
  True if Microsoft Word prints even pages in ascending order during manual duplex printing. Read/write Boolean.
- `DefaultBorderColorIndex As WdColorIndex  (read/write)`  
  Returns or sets the default line color for borders. Read/write WdColorIndex.
- `EnableMisusedWordsDictionary As Boolean  (read/write)`  
  True if Microsoft Word checks for misused words when checking the spelling and grammar in a document. Read/write Boolean.
- `AllowCombinedAuxiliaryForms As Boolean  (read/write)`  
  True if Microsoft Word ignores auxiliary verb forms when checking spelling in a Korean language document. Read/write Boolean.
- `HangulHanjaFastConversion As Boolean  (read/write)`  
  True if Microsoft Word automatically converts a word with only one suggestion during conversion between Hangul and Hanja. Read/write Boolean.
- `CheckHangulEndings As Boolean  (read/write)`  
  True if Microsoft Word automatically detects Hangul endings and ignores them during conversion from Hangul to Hanja. Read/write Boolean.
- `EnableHangulHanjaRecentOrdering As Boolean  (read/write)`  
  True if Microsoft Word displays the most recently used words at the top of the suggestions list during conversion between Hangul and Hanja. Read/write Boolean.
- `MultipleWordConversionsMode As WdMultipleWordConversionsMode  (read/write)`  
  Returns or sets the direction for conversion between Hangul and Hanja. Read/write WdMultipleWordConversionsMode.
- `DefaultBorderColor As WdColor  (read/write)`  
  Returns or sets the default 24-bit color to use for new Border objects. Read/write.
- `AllowPixelUnits As Boolean  (read/write)`  
  True if Microsoft Word uses pixels as the default unit of measurement for HTML features that support measurements. Read/write Boolean.
- `UseCharacterUnit As Boolean  (read/write)`  
  True if Microsoft Word uses characters as the default measurement unit for the current document. Read/write Boolean.
- `AllowCompoundNounProcessing As Boolean  (read/write)`  
  True if Microsoft Word ignores compound nouns when checking spelling in a Korean language document. Read/write Boolean.
- `AutoKeyboardSwitching As Boolean  (read/write)`  
  True if Microsoft Word automatically switches the keyboard language to match what you are typing at any given time. Read/write Boolean.
- `DocumentViewDirection As WdDocumentViewDirection  (read/write)`  
  Returns or sets the alignment and reading order for the entire document. Read/write WdDocumentViewDirection.
- `ArabicNumeral As WdArabicNumeral  (read/write)`  
  Returns or sets the numeral style for an Arabic language document. Read/write WdArabicNumeral.
- `MonthNames As WdMonthNames  (read/write)`  
  Returns or sets the direction for conversion between Hangul and Hanja. Read/write WdMonthNames.
- `CursorMovement As WdCursorMovement  (read/write)`  
  Returns or sets how the insertion point progresses within bidirectional text. Read/write WdCursorMovement.
- `VisualSelection As WdVisualSelection  (read/write)`  
  Returns or sets the selection behavior based on visual cursor movement in a right-to-left language document. Read/write WdVisualSelection.
- `ShowDiacritics As Boolean  (read/write)`  
  True if diacritics are visible in a right-to-left language document. Read/write Boolean.
- `ShowControlCharacters As Boolean  (read/write)`  
  True if bidirectional control characters are visible in the current document. Read/write Boolean.
- `AddControlCharacters As Boolean  (read/write)`  
  True if Microsoft Word adds bidirectional control characters when cutting and copying text. Read/write Boolean.
- `AddBiDirectionalMarksWhenSavingTextFile As Boolean  (read/write)`  
  True if Microsoft Word adds bidirectional control characters when saving a document as a text file. Read/write Boolean.
- `StrictInitialAlefHamza As Boolean  (read/write)`  
  True if the spelling checker uses spelling rules regarding Arabic words beginning with an alef hamza. Read/write Boolean.
- `StrictFinalYaa As Boolean  (read/write)`  
  True if the spelling checker uses spelling rules regarding Arabic words ending with the letter yaa. Read/write Boolean.
- `HebrewMode As WdHebSpellStart  (read/write)`  
  Returns or sets the mode for the Hebrew spelling checker. Read/write WdHebSpellStart.
- `ArabicMode As WdAraSpeller  (read/write)`  
  Returns or sets the mode for the Arabic spelling checker. Read/write WdAraSpeller.
- `AllowClickAndTypeMouse As Boolean  (read/write)`  
  True if Click and Type functionality is enabled. Read/write Boolean.
- `UseGermanSpellingReform As Boolean  (read/write)`  
  True if Microsoft Word uses the German post-reform spelling rules when checking spelling. Read/write Boolean.
- `InterpretHighAnsi As WdHighAnsiText  (read/write)`  
  Returns or sets the high-ANSI text interpretation behavior. Read/write WdHighAnsiText.
- `AddHebDoubleQuote As Boolean  (read/write)`  
  True if Microsoft Word encloses number formats in double quotation marks ("). Read/write Boolean.
- `UseDiffDiacColor As Boolean  (read/write)`  
  True if you can set the color of diacritics in the current document. Read/write Boolean.
- `DiacriticColorVal As WdColor  (read/write)`  
  Returns or sets the 24-bit color to be used for diacritics in a right-to-left language document. Read/write.
- `OptimizeForWord97byDefault As Boolean  (read/write)`  
  True if Microsoft Word optimizes all new documents for viewing in Word 97 by disabling any incompatible formatting. Read/write Boolean.
- `LocalNetworkFile As Boolean  (read/write)`  
  True if Microsoft Word creates a local copy of a file on the user's computer when editing a file stored on a network server. Read/write Boolean.
- `TypeNReplace As Boolean  (read/write)`  
  True for Microsoft Word to replace illegal South Asian characters. Read/write Boolean.
- `SequenceCheck As Boolean  (read/write)`  
  True to check the sequence of independent characters for South Asian text. Read/write Boolean.
- `DisableFeaturesbyDefault As Boolean  (read/write)`  
  True for Microsoft Word to disable in all documents all features introduced after the version of Word specified in the DisableFeaturesIntroducedAfterbyDefault. The default value is False. Read/write Boolean.
- `PasteAdjustWordSpacing As Boolean  (read/write)`  
  True if Microsoft Word automatically adjusts the spacing of words when cutting and pasting selections. Read/write Boolean.
- `PasteAdjustParagraphSpacing As Boolean  (read/write)`  
  True if Microsoft Word automatically adjusts the spacing of paragraphs when cutting and pasting selections. Read/write Boolean.
- `PasteAdjustTableFormatting As Boolean  (read/write)`  
  True if Microsoft Word automatically adjusts the formatting of tables when cutting and pasting selections. Read/write Boolean.
- `PasteSmartStyleBehavior As Boolean  (read/write)`  
  True if Microsoft Word intelligently merges styles when pasting a selection from a different document. Read/write Boolean.
- `PasteMergeFromPPT As Boolean  (read/write)`  
  True to merge text formatting when pasting from Microsoft PowerPoint. Read/write Boolean.
- `PasteMergeFromXL As Boolean  (read/write)`  
  True to merge table formatting when pasting from Microsoft Excel. Read/write Boolean.
- `CtrlClickHyperlinkToOpen As Boolean  (read/write)`  
  True if Microsoft Word requires holding down the Ctrl key while clicking to open a hyperlink. Read/write Boolean.
- `PictureWrapType As WdWrapTypeMerged  (read/write)`  
  Sets or returns a WdWrapTypeMerged that indicates how Microsoft Word wraps text around pictures. Read/write.
- `DisableFeaturesIntroducedAfterbyDefault As WdDisableFeaturesIntroducedAfter  (read/write)`  
  Disables all features introduced after a the specified version for all documents. Read/write WdDisableFeaturesIntroducedAfter.
- `PasteSmartCutPaste As Boolean  (read/write)`  
  True if Microsoft Word intelligently pastes selections into a document. Read/write Boolean.
- `DisplayPasteOptions As Boolean  (read/write)`  
  True for Microsoft Word to display the Paste Options button, which displays directly under newly pasted text. Read/write Boolean.
- `PromptUpdateStyle As Boolean  (read/write)`  
  True displays a message asking the user to verify whether they want to reformat a style or reapply the original style formatting when changing the formatting of styles. Read/write Boolean.
- `DefaultEPostageApp As String  (read/write)`  
  Sets or returns a String that represents the path and file name of the default electronic postage application. Read/write.
- `DefaultTextEncoding As MsoEncoding  (read/write)`  
  Returns or sets an MsoEncoding constant representing the code page, or character set, that Microsoft Word uses for all documents saved as encoded text files. Read/write.
- `WarnBeforeSavingPrintingSendingMarkup As Boolean  (read/write)`  
  True for Microsoft Word to display a warning when saving, printing, or sending as email a document containing comments or tracked changes. Read/write Boolean.
- `StoreRSIDOnSave As Boolean  (read/write)`  
  True for Microsoft Word to assign a random number to changes in a document, each time a document is saved, to facilitate comparing and merging documents. Read/write Boolean.
- `ShowFormatError As Boolean  (read/write)`  
  True for Microsoft Word to mark inconsistencies in formatting by placing a squiggly underline beneath text formatted similarly to other formatting that is used more frequently in a document. Read/write Boolean.
- `FormatScanning As Boolean  (read/write)`  
  True for Microsoft Word to keep track of all formatting in a document. Read/write Boolean.
- `PasteMergeLists As Boolean  (read/write)`  
  True to merge the formatting of pasted lists with surrounding lists. Read/write Boolean.
- `AutoCreateNewDrawings As Boolean  (read/write)`  
  True for Microsoft Word to draw newly created shapes in a drawing canvas. Read/write Boolean.
- `SmartParaSelection As Boolean  (read/write)`  
  True for Microsoft Word to include the paragraph mark in a selection when selecting most or all of a paragraph. Read/write Boolean.
- `RevisionsBalloonPrintOrientation As WdRevisionsBalloonPrintOrientation  (read/write)`  
  Returns or sets a WdRevisionsBalloonPrintOrientation constant that represents the direction of revision and comment balloons when they are printed. Read/write.
- `CommentsColor As WdColorIndex  (read/write)`  
  Returns or sets a WdColorIndex constant that represents the color of comments in a document. Read/write.
- `PrintXMLTag As Boolean  (read/write)`  
  Returns a Boolean that represents whether to print the XML tags when printing a document. Corresponds to the XML tags check box on the Print tab in the Options dialog box. .
- `PrintBackgrounds As Boolean  (read/write)`  
  Returns a Boolean that represents whether background colors and images are printed when a document is printed.
- `AllowReadingMode As Boolean  (read/write)`  
  True indicates that Microsoft Word opens documents in Reading Layout view. Read/write Boolean.
- `ShowMarkupOpenSave As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Microsoft Word displays hidden markup when opening or saving a file.
- `SmartCursoring As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether smart cursoring is enabled. True enables smart cursoring.
- `MoveToTextMark As WdMoveToTextMark  (read/write)`  
  Returns or sets a WdMoveToTextMark constant that represents the type of revision mark to use for moved text. Read/write.
- `MoveFromTextMark As WdMoveFromTextMark  (read/write)`  
  Returns or sets a WdMoveFromTextMark constant that represents the type of revision mark to use for moved text. Read/write.
- `BibliographyStyle As String  (read/write)`  
  Returns or sets a String that represents the name of the style to use for formatting bibliographies. Read/write.
- `BibliographySort As String  (read/write)`  
  Returns or sets a String that represents the order in which to display sources in the Source Manager dialog box. Read/write.
- `InsertedCellColor As WdCellColor  (read/write)`  
  Returns or sets a WdCellColor constant that represents the color for an inserted table cell. Read/write.
- `DeletedCellColor As WdCellColor  (read/write)`  
  Returns or sets a WdCellColor constant that represents the color for a deleted cell. Read/write.
- `MergedCellColor As WdCellColor  (read/write)`  
  Returns or sets a WdCellColor constant that represents the color for merged table cells. Read/write.
- `SplitCellColor As WdCellColor  (read/write)`  
  Returns or sets a WdCellColor that represents the color for split table cells. Read/write.
- `ShowSelectionFloaties As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether mini toolbars display when a user selects text. Read/write.
- `ShowMenuFloaties As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to display mini toolbars when the user right-clicks in the document window. Read/write.
- `ShowDevTools As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the Developer tab is displayed in the ribbon. Read/write.
- `EnableLivePreview As Boolean  (read/write)`  
  Sets or returns a Boolean that represents whether to show or hide gallery previews that appear when using galleries that support previewing. True shows a preview in your document before applying the command. Read/write.
- `OMathAutoBuildUp As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Microsoft Word automatically converts equations to professional format. True indicates that Word automatically converts equations to professional format. Read/write.
- `PasteFormatWithinDocument As WdPasteOptions  (read/write)`  
  Returns or sets a WdPasteOptions constant that represents how text is pasted when text is copied or cut and then pasted in the same document. Read/write.
- `PasteFormatBetweenDocuments As WdPasteOptions  (read/write)`  
  Returns or sets a WdPasteOptions constant that represents how text is pasted when text is copied from another Microsoft Office Word document. Read/write.
- `PasteFormatBetweenStyledDocuments As WdPasteOptions  (read/write)`  
  Returns or sets a WdPasteOptions constant that represents how text is pasted when text is copied from a document that uses styles. Read/write.
- `PasteFormatFromExternalSource As WdPasteOptions  (read/write)`  
  Returns or sets a WdPasteOptions constant that represents how text is pasted when text is copied from an external source, such as a webpage. Read/write.
- `PasteOptionKeepBulletsAndNumbers As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to keep bullets and numbering when selecting Keep text only from the Paste Options context menu. Read/write.
- `INSKeyForOvertype As Boolean  (read/write)`  
  True if the INS key can be used for switching Overtype on and off. Read/write Boolean.
- `RepeatWord As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to mark words that are repeated when spelling is checked. True flags repeated words. Read/write.
- `FrenchReform As WdFrenchSpeller  (read/write)`  
  Returns or sets a WdFrenchSpeller constant that represents which spelling dictionary to use for regions of text with language formatting set to French. Read/write.
- `ContextualSpeller As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to use the contextual speller to check spelling based on the context of a word and the words around it. Read/write.
- `MoveToTextColor As WdColorIndex  (read/write)`  
  Returns or sets a WdColorIndex constant that represents the color of moved text. Read/write.
- `MoveFromTextColor As WdColorIndex  (read/write)`  
  Returns or sets a WdColorIndex constant that represents the color of moved text. Read/write.
- `OMathCopyLF As Boolean  (read/write)`  
  Returns or sets a Boolean that represents how equations are represented in plain text. True indicates equations are represented in Linear Format. False indicates equations are represented in MathML. Read/write.
- `UseNormalStyleForList As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Microsoft Word uses the Normal style for bullets and numbering. Read/write.
- `AllowOpenInDraftView As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to allow users to open documents in draft view. Read/write.
- `EnableLegacyIMEMode As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to enable legacy IME mode. Read/write.
- `DoNotPromptForConvert As Boolean  (read/write)`  
  Sets or returns a Boolean that represents whether to prompt a warning dialog when the Convert command is invoked for documents that are in compatibility mode. Read/write.
- `PrecisePositioning As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Word optimizes character positioning for print layout rather than on-screen readability. True disables the default setting that compresses character spacing to facilitate on-screen readability and enables character spacing for print media. Read/write.
- `UpdateStyleListBehavior As WdUpdateStyleListBehavior  (read/write)`  
  Returns or sets a WdUpdateStyleListBehavior constant that specifies the behavior Word should take when updating a style to match a selection that contains numbering or bullets. Read/write.
- `StrictTaaMarboota As Boolean  (read/write)`  
  True if the spelling checker uses spelling rules to flag Arabic words ending with haa instead of taa marboota. Read/write.
- `StrictRussianE As Boolean  (read/write)`  
  True if the spelling checker uses spelling rules regarding Russian words that use the strict  character. Read/write.
- `SpanishMode As WdSpanishSpeller  (read/write)`  
  Returns or sets the mode for the Spanish speller. Read/write WdSpanishSpeller.
- `PortugalReform As WdPortugueseReform  (read/write)`  
  Returns or sets the mode for the European Portuguese speller. Read/write WdPortugueseReform.
- `BrazilReform As WdPortugueseReform  (read/write)`  
  Returns or sets the mode for the Brazilian Portuguese speller. Read/write WdPortugueseReform.
- `UpdateFieldsWithTrackedChangesAtPrint As Boolean  (read/write)`  
  True if Word allows fields containing tracked changes to update before printing. Read/write.
- `DisplayAlignmentGuides As Boolean  (read/write)`  
  Returns or sets a Boolean that specifies whether alignment guides are enabled in the user interface. Read/write.
- `PageAlignmentGuides As Boolean  (read/write)`  
  Returns or sets a Boolean that specifies whether page alignment guides are displayed in the user interface. Read/write.
- `MarginAlignmentGuides As Boolean  (read/write)`  
  Returns or sets a Boolean that specifies whether margin alignment guides are displayed in the user interface. Read/write.
- `ParagraphAlignmentGuides As Boolean  (read/write)`  
  Returns or sets a Boolean that specifies whether paragraph alignment guides are displayed in the user interface. Read/write.
- `EnableLiveDrag As Boolean  (read/write)`  
  Returns or sets a Boolean that is True if live drag is enabled. Read/write.
- `UseSubPixelPositioning As Boolean  (read/write)`  
  Returns or sets a Boolean that specifies whether sub-pixel positioning is enabled. Read/write.
- `AlertIfNotDefault As Boolean  (read/write)`  
  Returns or sets a Boolean that is True if users are notified if Word is not the default program for viewing and editing documents. Read/write.
- `EnableProofingToolsAdvertisement As Boolean  (read/write)`  
  Returns or sets a Boolean value that, when True, specifies that users be notified when additional proofing tools are available for download. Read/write.
- `PreferCloudSaveLocations As Boolean  (read/write)`  
  True to save new documents in web locations by default. Read/write Boolean.
- `ExpandHeadingsOnOpen As Boolean  (read/write)`  
  True to expand all headings in the document when the document opens. Read/write Boolean.
- `UseLocalUserInfo As Boolean  (read/write)`  
  Returns or sets a Boolean; True if Microsoft Word identifies the document author based upon the User name and Initials settings on the General tab of the Options dialog box, and False if Word identifies the author based on the account information with which the user signed in to Office. Read/write.
- `CloudSignInOption As Boolean  (read/write)`  
  True to give users the option to sign in to Microsoft OneDrive and other cloud locations. Read/write Boolean.
- `ShowPopupAddRowColToTable As Boolean  (read/write)`
- `LiveWordCount As Boolean  (read/write)`
- `AllowCoAuthoringOnFilesWithMacros As Boolean  (read/write)`
