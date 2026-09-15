# Find

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209B0-0000-0000-C000-000000000046}  

Represents the criteria for a find operation.

**Remarks:** The properties and methods of the Find object correspond to the options in the Find and Replace dialog box. Use the Find property to return a Find object. The following example finds and selects the next occurrence of the word "hi." The following example finds all occurrences of the word "hi" in the active document and replaces the word with "hello." If you've gotten to the Find object from the Selection object, the selection is changed when text matching the find criteria is found. The following example selects the next occurrence of the word "blue." If you've gotten to the Find object from the Range object, the selection isn't changed when text matching the find criteria is found, but the Range object is redefined. The following example locates the first occurrence of the word "blue" in the active document. If "blue" is found in the document, myRange is redefined and bold formatting is applied to "blue."

## Properties (36)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Find object.
- `Forward As Boolean  (read/write)`  
  True if the find operation searches forward through the document. Read/write Boolean.
- `Font As Font  (read/write)`  
  Returns or sets a Font object that represents the character formatting of the specified object. Read/write Font.
- `Found As Boolean  (read-only)`  
  True if the search produces a match. Read-only Boolean.
- `MatchAllWordForms As Boolean  (read/write)`  
  True if all forms of the text to find are found by the find operation (for instance, if the text to find is "sit," "sat" and "sitting" are found as well). Read/write Boolean.
- `MatchCase As Boolean  (read/write)`  
  True if the find operation is case-sensitive. The default is False. Read/write Boolean.
- `MatchWildcards As Boolean  (read/write)`  
  True if the text to find contains wildcards. Read/write Boolean.
- `MatchSoundsLike As Boolean  (read/write)`  
  True if words that sound similar to the text to find are returned by the find operation. Read/write Boolean.
- `MatchWholeWord As Boolean  (read/write)`  
  True if the find operation locates only entire words and not text that's part of a larger word. Read/write Boolean.
- `MatchFuzzy As Boolean  (read/write)`  
  True if Microsoft Word uses the nonspecific search options for Japanese text during a search. Read/write Boolean.
- `MatchByte As Boolean  (read/write)`  
  True if Microsoft Word distinguishes between full-width and half-width letters or characters during a search. Read/write Boolean.
- `ParagraphFormat As ParagraphFormat  (read/write)`  
  Returns or sets a ParagraphFormat object that represents the paragraph settings for the specified find operation. Read/write.
- `Style As Variant  (read/write)`  
  Returns or sets the style for the specified object. Read/write Variant.
- `Text As String  (read/write)`  
  Returns or sets the text to find. Read/write String.
- `LanguageID As WdLanguageID  (read/write)`  
  Returns or sets the language for the specified Find object. Read/write WdLanguageID.
- `Highlight As Long  (read/write)`  
  True if highlight formatting is included in the find criteria. Read/write Long.
- `Replacement As Replacement  (read-only)`  
  Returns a Replacement object that contains the criteria for a replace operation.
- `Frame As Frame  (read-only)`  
  Returns a Frame object that represents the frame formatting for the specified style or find-and-replace operation. Read-only.
- `Wrap As WdFindWrap  (read/write)`  
  Returns or sets what happens if the search begins at a point other than the beginning of the document and the end of the document is reached (or vice versa if Forward is set to False) or if the search text isn't found in the specified selection or range. Read/write WdFindWrap.
- `Format As Boolean  (read/write)`  
  True if formatting is included in the find operation. Read/write Boolean.
- `LanguageIDFarEast As WdLanguageID  (read/write)`  
  Returns or sets an East Asian language for the specified object. Read/write WdLanguageID.
- `LanguageIDOther As WdLanguageID  (read/write)`  
  Returns or sets the language for the specified object. Read/write WdLanguageID.
- `CorrectHangulEndings As Boolean  (read/write)`  
  True if Microsoft Word automatically corrects Hangul endings when replacing Hangul text. Read/write Boolean.
- `NoProofing As Long  (read/write)`  
  True if Microsoft Word finds or replaces text that the spelling and grammar checker ignores. Read/write Long.
- `MatchKashida As Boolean  (read/write)`  
  True if find operations match text with matching kashidas in an Arabic language document. Read/write Boolean.
- `MatchDiacritics As Boolean  (read/write)`  
  True if find operations match text with matching diacritics in a right-to-left language document. Read/write Boolean.
- `MatchAlefHamza As Boolean  (read/write)`  
  True if find operations match text with matching alef hamzas in an Arabic language document. Read/write Boolean.
- `MatchControl As Boolean  (read/write)`  
  True if find operations match text with matching bidirectional control characters in a right-to-left language document. Read/write Boolean.
- `MatchPhrase As Boolean  (read/write)`  
  True ignores all white space and control characters between words. Read/write.
- `MatchPrefix As Boolean  (read/write)`  
  True to match words beginning with the search string. Read/write.
- `MatchSuffix As Boolean  (read/write)`  
  True to match words ending with the search string. Read/write.
- `IgnoreSpace As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether a find operation should ignore extra white space in found text. Read/write.
- `IgnorePunct As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether a find operation should ignore punctuation in found text. Read/write.
- `HanjaPhoneticHangul As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to locate phonetic Hangul and hanja characters in a Korean language find operation. Read/write.

## Methods (7)

- `ClearFormatting()`  
  Removes text and paragraph formatting from the text specified in a find or replace operation.
- `SetAllFuzzyOptions()`  
  Activates all nonspecific search options associated with Japanese text.
- `ClearAllFuzzyOptions()`  
  Clears all nonspecific search options associated with Japanese text.
- `Execute([FindText As Variant], [MatchCase As Variant], [MatchWholeWord As Variant], [MatchWildcards As Variant], [MatchSoundsLike As Variant], [MatchAllWordForms As Variant], [Forward As Variant], [Wrap As Variant], [Format As Variant], [ReplaceWith As Variant], [Replace As Variant], [MatchKashida As Variant], [MatchDiacritics As Variant], [MatchAlefHamza As Variant], [MatchControl As Variant]) As Boolean`  
  Runs the specified find operation. Returns True if the find operation is successful. Boolean.
    - `FindText As Variant` (optional): The text to be searched for. Use an empty string ("") to search for formatting only. You can search for special characters by specifying appropriate character codes. For example, "^p" corresponds to a paragraph mark and "^t" corresponds to a tab character.
    - `MatchCase As Variant` (optional): True to specify that the find text be case-sensitive. Corresponds to the Match case check box in the Find and Replace dialog box (Edit menu).
    - `MatchWholeWord As Variant` (optional): True to have the find operation locate only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
    - `MatchWildcards As Variant` (optional): True to have the find text be a special search operator. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
    - `MatchSoundsLike As Variant` (optional): True to have the find operation locate words that sound similar to the find text. Corresponds to the Sounds like check box in the Find and Replace dialog box.
    - `MatchAllWordForms As Variant` (optional): True to have the find operation locate all forms of the find text (for example, "sit" locates "sitting" and "sat"). Corresponds to the Find all word forms check box in the Find and Replace dialog box.
    - `Forward As Variant` (optional): True to search forward (toward the end of the document).
    - `Wrap As Variant` (optional): Controls what happens if the search begins at a point other than the beginning of the document and the end of the document is reached (or vice versa if Forward is set to False). This argument also controls what happens if there is a selection or range and the search text is not found in the selection or range. Can be one of the WdFindWrap constants.
    - `Format As Variant` (optional): True to have the find operation locate formatting in addition to, or instead of, the find text.
    - `ReplaceWith As Variant` (optional): The replacement text. To delete the text specified by the Find argument, use an empty string (""). You specify special characters and advanced search criteria just as you do for the Find argument. To specify a graphic object or other nontext item as the replacement, move the item to the Clipboard and specify "^c" for ReplaceWith.
    - `Replace As Variant` (optional): Specifies how many replacements are to be made: one, all, or none. Can be any WdReplace constant.
    - `MatchKashida As Variant` (optional): True if find operations match text with matching kashidas in an Arabic-language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `MatchDiacritics As Variant` (optional): True if find operations match text with matching diacritics in a right-to-left language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `MatchAlefHamza As Variant` (optional): True if find operations match text with matching alef hamzas in an Arabic-language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `MatchControl As Variant` (optional): True if find operations match text with matching bidirectional control characters in a right-to-left language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
- `HitHighlight(FindText As Variant, [HighlightColor As Variant], [TextColor As Variant], [MatchCase As Variant], [MatchWholeWord As Variant], [MatchPrefix As Variant], [MatchSuffix As Variant], [MatchPhrase As Variant], [MatchWildcards As Variant], [MatchSoundsLike As Variant], [MatchAllWordForms As Variant], [MatchByte As Variant], [MatchFuzzy As Variant], [MatchKashida As Variant], [MatchDiacritics As Variant], [MatchAlefHamza As Variant], [MatchControl As Variant], [IgnoreSpace As Variant], [IgnorePunct As Variant], [HanjaPhoneticHangul As Variant]) As Boolean`  
  Highlights all found matches and returns a Boolean that represents whether matches were found.
    - `FindText As Variant` (required): Specifies the text to find. Use an empty string ("") to search for formatting only. You can search for special characters by specifying appropriate character codes. For example, "^p" corresponds to a paragraph mark and "^t" corresponds to a tab character.
    - `HighlightColor As Variant` (optional): Specifies the highlight color for the text. Can be any RGB color or one of the WdColor constants.
    - `TextColor As Variant` (optional): Specifies the color of the text. Can be any RGB color or one of the WdColor constants.
    - `MatchCase As Variant` (optional): True to specify that the find text be case-sensitive. Corresponds to the Match case check box in the Find and Replace dialog box.
    - `MatchWholeWord As Variant` (optional): True to have the find operation locate only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
    - `MatchPrefix As Variant` (optional): True to match words beginning with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
    - `MatchSuffix As Variant` (optional): True to match words ending with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
    - `MatchPhrase As Variant` (optional): True ignores all white space and control characters between words.
    - `MatchWildcards As Variant` (optional): True to have the find text be a special search operator. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
    - `MatchSoundsLike As Variant` (optional): True to have the find operation locate words that sound similar to the find text. Corresponds to the Sounds like check box in the Find and Replace dialog box.
    - `MatchAllWordForms As Variant` (optional): True to have the find operation locate all forms of the find text (for example, "sit" locates "sitting" and "sat"). Corresponds to the Find all word forms check box in the Find and Replace dialog box.
    - `MatchByte As Variant` (optional): True to distinguish between full-width and half-width letters or characters during a search.
    - `MatchFuzzy As Variant` (optional): True to use the nonspecific search options for Japanese text during a search. Read/write.
    - `MatchKashida As Variant` (optional): True if find operations match text with matching kashidas in an Arabic-language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `MatchDiacritics As Variant` (optional): True if find operations match text with matching diacritics in a right-to-left language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `MatchAlefHamza As Variant` (optional): True if find operations match text with matching alef hamzas in an Arabic-language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `MatchControl As Variant` (optional): True if find operations match text with matching bidirectional control characters in a right-to-left language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreSpace As Variant` (optional): True ignores all white space between words. Corresponds to the Ignore white-space characters check box in the Find and Replace dialog box.
    - `IgnorePunct As Variant` (optional): True ignores all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
    - `HanjaPhoneticHangul As Variant` (optional): True ignores phonetic hangul and hanja characters. Available only if you have support for Korean languages.
- `ClearHitHighlight() As Boolean`  
  Removes the highlighting for all text located in a hit highlighting find operation, and returns a Boolean that represents whether the operation was successful.
- `Execute2007([FindText As Variant], [MatchCase As Variant], [MatchWholeWord As Variant], [MatchWildcards As Variant], [MatchSoundsLike As Variant], [MatchAllWordForms As Variant], [Forward As Variant], [Wrap As Variant], [Format As Variant], [ReplaceWith As Variant], [Replace As Variant], [MatchKashida As Variant], [MatchDiacritics As Variant], [MatchAlefHamza As Variant], [MatchControl As Variant], [MatchPrefix As Variant], [MatchSuffix As Variant], [MatchPhrase As Variant], [IgnoreSpace As Variant], [IgnorePunct As Variant]) As Boolean`  
  Runs the specified find operation. Returns True if the find operation is successful.
    - `FindText As Variant` (optional): The text to be searched for. Use an empty string ("") to search for formatting only. You can search for special characters by specifying appropriate character codes. For example, "^p" corresponds to a paragraph mark and "^t" corresponds to a tab character.
    - `MatchCase As Variant` (optional): True to specify that the text to find should be case-sensitive. Corresponds to the Match case check box in the Find and Replace dialog box (on the Home tab in the Editing group).
    - `MatchWholeWord As Variant` (optional): True to find only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
    - `MatchWildcards As Variant` (optional): True to use wildcard search operators in the text to find. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
    - `MatchSoundsLike As Variant` (optional): True to locate words that sound similar to the text to find. Corresponds to the Sounds like check box in the Find and Replace dialog box.
    - `MatchAllWordForms As Variant` (optional): True to locate all forms of the text to find (for example, "sit" locates "sitting" and "sat"). Corresponds to the Find all word forms check box in the Find and Replace dialog box.
    - `Forward As Variant` (optional): True to search forward (toward the end of the document).
    - `Wrap As Variant` (optional): One of the WdFindWrap constants that controls what happens if the search begins at a point other than the beginning of the document and the end of the document is reached (or vice versa if Forward is set to False). This argument also controls what happens if there is a selection or range and the search text is not found in the selection or range.
    - `Format As Variant` (optional): True to locate formatting in addition to, or instead of, the text to find.
    - `ReplaceWith As Variant` (optional): The replacement text. To delete the text specified by the Find argument, use an empty string (""). You specify special characters and advanced search criteria just as you do for the Find argument. To specify a graphic object or other nontext item as the replacement, move the item to the Clipboard and specify "^c" for ReplaceWith.
    - `Replace As Variant` (optional): One of the WdReplace constants that specifies how many replacements are to be made: one, all, or none.
    - `MatchKashida As Variant` (optional): True to find matching kashidas in an Arabic-language document. This argument may not be available to you, depending on the language support (for example, U.S. English) that you have selected or installed.
    - `MatchDiacritics As Variant` (optional): True to find matching diacritics in a right-to-left language document. This argument may not be available to you, depending on the language support (for example, U.S. English) that you have selected or installed.
    - `MatchAlefHamza As Variant` (optional): True to find matching alef hamzas in an Arabic-language document. This argument may not be available to you, depending on the language support (for example, U.S. English) that you have selected or installed.
    - `MatchControl As Variant` (optional): True to find matching bidirectional control characters in a right-to-left language document. This argument may not be available to you, depending on the language support (for example, U.S. English) that you have selected or installed.
    - `MatchPrefix As Variant` (optional): True to find words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
    - `MatchSuffix As Variant` (optional): True to find words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
    - `MatchPhrase As Variant` (optional): True to ignore all white space and control characters between words.
    - `IgnoreSpace As Variant` (optional): True to ignore all white space between words. Corresponds to the Ignore white-space characters check box in the Find and Replace dialog box.
    - `IgnorePunct As Variant` (optional): True to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
