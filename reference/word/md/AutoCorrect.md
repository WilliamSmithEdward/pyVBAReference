# AutoCorrect

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020949-0000-0000-C000-000000000046}  

Represents the AutoCorrect functionality in Word.

**Remarks:** Use the AutoCorrect property to return the AutoCorrect object. The following example enables the AutoCorrect options and creates an AutoCorrect entry. The Entries property returns the Entries object that represents the AutoCorrect entries in the AutoCorrect dialog box.

## Properties (22)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified AutoCorrect object.
- `CorrectDays As Boolean  (read/write)`  
  True if Word automatically capitalizes the first letter of days of the week. Read/write Boolean.
- `CorrectInitialCaps As Boolean  (read/write)`  
  True if Word automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, "WOrd" is corrected to "Word." Read/write Boolean.
- `CorrectSentenceCaps As Boolean  (read/write)`  
  True if Word automatically capitalizes the first letter in each sentence. Read/write Boolean.
- `ReplaceText As Boolean  (read/write)`  
  True if Microsoft Word automatically replaces specified text with entries from the AutoCorrect list. Read/write Boolean.
- `Entries As AutoCorrectEntries  (read-only)`  
  Returns an AutoCorrectEntries collection that represents the current list of AutoCorrect entries.
- `FirstLetterExceptions As FirstLetterExceptions  (read-only)`  
  Returns a FirstLetterExceptions collection that represents the list of abbreviations after which Word won't automatically capitalize the next letter. Read-only.
- `FirstLetterAutoAdd As Boolean  (read/write)`  
  True if Word automatically adds abbreviations to the list of AutoCorrect First Letter exceptions. Read/write Boolean.
- `TwoInitialCapsExceptions As TwoInitialCapsExceptions  (read-only)`  
  Returns a TwoInitialCapsExceptions collection that represents the list of terms containing mixed capitalization that Word won't correct automatically.
- `TwoInitialCapsAutoAdd As Boolean  (read/write)`  
  True if Microsoft Word automatically adds words to the list of AutoCorrect Initial Caps exceptions. A word is added to this list if you delete and then retype the uppercase letter (following the initial uppercase letter) that Word changed to lowercase. Read/write Boolean.
- `CorrectCapsLock As Boolean  (read/write)`  
  True if Word automatically corrects instances in which you use the CAPS LOCK key inadvertently as you type. Read/write Boolean.
- `CorrectHangulAndAlphabet As Boolean  (read/write)`  
  True if Microsoft Word automatically applies the correct font to Latin words typed in the middle of Hangul text or vice versa. Read/write Boolean.
- `HangulAndAlphabetExceptions As HangulAndAlphabetExceptions  (read-only)`  
  Returns a HangulAndAlphabetExceptions collection that represents the list of Hangul and alphabet AutoCorrect exceptions.
- `HangulAndAlphabetAutoAdd As Boolean  (read/write)`  
  True if Microsoft Word automatically adds words to the list of Hangul and alphabet AutoCorrect exceptions. Read/write Boolean.
- `ReplaceTextFromSpellingChecker As Boolean  (read/write)`  
  True if Microsoft Word automatically replaces misspelled text with suggestions from the spelling checker as the user types. Word only replaces words that contain a single misspelling and for which the spelling dictionary only lists one alternative. Read/write Boolean.
- `OtherCorrectionsAutoAdd As Boolean  (read/write)`  
  True if Microsoft Word automatically adds words to the list of AutoCorrect exceptions on the Other Corrections tab in the AutoCorrect Exceptions dialog box (AutoCorrect Options command, Tools menu). Word adds a word to this list if you delete and then retype a word that you didn't want Word to correct. Read/write Boolean.
- `OtherCorrectionsExceptions As OtherCorrectionsExceptions  (read-only)`  
  Returns an OtherCorrectionsExceptions collection that represents the list of words that Microsoft Word won't correct automatically.
- `CorrectKeyboardSetting As Boolean  (read/write)`  
  True if Microsoft Word automatically transposes words to their native alphabet if you type text in a language other than the current keyboard language. Read/write Boolean.
- `CorrectTableCells As Boolean  (read/write)`  
  True to automatically capitalize the first letter of table cells. Read/write Boolean.
- `DisplayAutoCorrectOptions As Boolean  (read/write)`  
  True for Microsoft Word to display the AutoCorrect Options button. Read/write Boolean.
