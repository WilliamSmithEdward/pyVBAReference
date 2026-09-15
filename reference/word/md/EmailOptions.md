# EmailOptions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209DB-0000-0000-C000-000000000046}  

Contains global application-level attributes used by Microsoft Word when you create and edit email messages and replies.

**Remarks:** Use the EmailOptions property to return the EmailOptions object. This example changes the font color of the default style used to compose new email messages. This example sets Word to mark comments in email messages with the initials "WK." This example changes the signatures Word appends to new outgoing email messages and email message replies.

## Properties (38)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified EmailOptions object.
- `UseThemeStyle As Boolean  (read/write)`  
  True if new email messages use the character style defined by the default email message theme. Read/write Boolean.
- `MarkCommentsWith As String  (read/write)`  
  Returns or sets the string with which Microsoft Word marks comments in email messages. Read/write String.
- `MarkComments As Boolean  (read/write)`  
  True if Microsoft Word marks the user's comments in email messages. Read/write Boolean.
- `EmailSignature As EmailSignature  (read-only)`  
  Returns an EmailSignature object that represents the signatures Microsoft Word appends to outgoing email messages. Read-only.
- `ComposeStyle As Style  (read-only)`  
  Returns a Style object that represents the style used to compose new email messages. Read-only.
- `ReplyStyle As Style  (read-only)`  
  Returns a Style object that represents the style used when replying to email messages.
- `ThemeName As String  (read/write)`  
  Returns or sets the name of the theme plus any theme formatting options to use for new email messages. Read/write String.
- `NewColorOnReply As Boolean  (read/write)`  
  True specifies whether a user needs to choose a new color for reply text when replying to email. Read/write Boolean.
- `PlainTextStyle As Style  (read-only)`  
  Returns the Style object that represents the text attributes for email messages that are sent or received using plain text.
- `UseThemeStyleOnReply As Boolean  (read/write)`  
  True for Microsoft Word to use a theme when replying to email. Read/write Boolean.
- `AutoFormatAsYouTypeApplyHeadings As Boolean  (read/write)`  
  True if styles are automatically applied to headings as you type. Read/write Boolean.
- `AutoFormatAsYouTypeApplyBorders As Boolean  (read/write)`  
  True if a series of three or more hyphens (-), equal signs (=), or underscore characters (_) are automatically replaced by a specific border line when the ENTER key is pressed. Read/write Boolean.
- `AutoFormatAsYouTypeApplyBulletedLists As Boolean  (read/write)`  
  True if bullet characters (such as asterisks, hyphens, and greater-than signs) are replaced with bullets. Read/write Boolean.
- `AutoFormatAsYouTypeApplyNumberedLists As Boolean  (read/write)`  
  True if paragraphs are automatically formatted as numbered lists. Read/write Boolean.
- `AutoFormatAsYouTypeReplaceQuotes As Boolean  (read/write)`  
  True if straight quotation marks are automatically changed to smart (curly) quotation marks as you type. Read/write Boolean.
- `AutoFormatAsYouTypeReplaceSymbols As Boolean  (read/write)`  
  True if two consecutive hyphens (--) are replaced with an en dash (-) or an em dash (-) as you type. Read/write Boolean.
- `AutoFormatAsYouTypeReplaceOrdinals As Boolean  (read/write)`  
  True if the ordinal number suffixes "st", "nd", "rd", and "th" are replaced with the same letters in superscript as you type; for example, "1st" is replaced with "1" followed by "st" formatted as superscript. Read/write Boolean.
- `AutoFormatAsYouTypeReplaceFractions As Boolean  (read/write)`  
  True if typed fractions are replaced with fractions from the current character set as you type; for example, "1/2" is replaced with "." Read/write Boolean.
- `AutoFormatAsYouTypeReplacePlainTextEmphasis As Boolean  (read/write)`  
  True if manual emphasis characters are automatically replaced with character formatting as you type; for example, "bold" is changed to " bold ". Read/write Boolean.
- `AutoFormatAsYouTypeFormatListItemBeginning As Boolean  (read/write)`  
  True if Word repeats character formatting applied to the beginning of a list item to the next list item. Read/write Boolean.
- `AutoFormatAsYouTypeDefineStyles As Boolean  (read/write)`  
  True if Word automatically creates new styles based on manual formatting. Read/write Boolean.
- `AutoFormatAsYouTypeReplaceHyperlinks As Boolean  (read/write)`  
  True if email addresses, server and share names (also known as UNC paths), and Internet addresses (also known as URLs) are automatically changed to hyperlinks as you type. Read/write Boolean.
- `AutoFormatAsYouTypeApplyTables As Boolean  (read/write)`  
  True if Word automatically creates a table when you type a plus sign, a series of hyphens, another plus sign, and so on, and then press ENTER. Read/write Boolean.
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
- `RelyOnCSS As Boolean  (read/write)`  
  True if cascading style sheets (CSS) are used for font formatting when you view a saved document in a web browser. Read/write Boolean.
- `HTMLFidelity As WdEmailHTMLFidelity  (read/write)`  
  Strips HTML tags used for opening HTML files in Word but not required for display. Read/write WdEmailHTMLFidelity.
- `TabIndentKey As Boolean  (read/write)`  
  True if the TAB and BACKSPACE keys can be used to increase and decrease, respectively, the left indent of paragraphs and if the BACKSPACE key can be used to change right-aligned paragraphs to centered paragraphs and centered paragraphs to left-aligned paragraphs. Read/write Boolean.
