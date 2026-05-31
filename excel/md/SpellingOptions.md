# SpellingOptions

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024465-0000-0000-C000-000000000046}  

Represents the various spell checking options for a worksheet.

**Remarks:** Use the SpellingOptions property of the Application object to return a SpellingOptions object. After a SpellingOptions object is returned, you can use the following properties to set or return various spell checking options: - ArabicModes - DictLang - GermanPostReform - HebrewModes - IgnoreCaps - IgnoreFileNames - IgnoreMixedDigits - KoreanCombineAux - KoreanProcessCompound - KoreanUseAutoChangeList - SuggestMainOnly - UserDict

**Example:**

```vba
Sub IgnoreAllCAPS()

 ' Place misspelled versions of the same word in all caps and mixed case.
 Range("A1").Formula = "Testt"
 Range("A2").Formula = "TESTT"

 With Application.SpellingOptions
 .SuggestMainOnly = True
 .IgnoreCaps = True
 End With

 ' Run a spell check.
 Cells.CheckSpelling

End Sub
```

## Properties (19)

- `DictLang As Long  (read/write)`  
  Selects the dictionary language used when Microsoft Excel performs spelling checks. Read/write Long.
- `UserDict As String  (read/write)`  
  Instructs Microsoft Excel to create a custom dictionary to which new words can be added when performing spelling checks on a worksheet. Read/write String.
- `IgnoreCaps As Boolean  (read/write)`  
  False instructs Microsoft Excel to check for uppercase words; True instructs Excel to ignore words in uppercase when using the spelling checker. Read/write Boolean.
- `SuggestMainOnly As Boolean  (read/write)`  
  When set to True, instructs Microsoft Excel to suggest words from only the main dictionary when using the spelling checker. False removes the limits of suggesting words from only the main dictionary when using the spelling checker. Read/write Boolean.
- `IgnoreMixedDigits As Boolean  (read/write)`  
  False instructs Microsoft Excel to check for mixed digits; True instructs Excel to ignore mixed digits when checking spelling. Read/write Boolean.
- `IgnoreFileNames As Boolean  (read/write)`  
  False instructs Microsoft Excel to check for Internet and file addresses; True instructs Excel to ignore Internet and file addresses when using the spell checker. Read/write Boolean.
- `GermanPostReform As Boolean  (read/write)`  
  True to check the spelling of words by using the German post-reform rules. False cancels this feature. Read/write Boolean.
- `KoreanCombineAux As Boolean  (read/write)`  
  When set to True, Microsoft Excel combines Korean auxiliary verbs and adjectives when spelling is checked. Read/write Boolean.
- `KoreanUseAutoChangeList As Boolean  (read/write)`  
  When set to True, this enables Microsoft Excel to use the auto-change list for Korean words when using the spelling checker. Read/write Boolean.
- `KoreanProcessCompound As Boolean  (read/write)`  
  When set to True, this enables Microsoft Excel to process Korean compound nouns when using the spelling checker. Read/write Boolean.
- `HebrewModes As XlHebrewModes  (read/write)`  
  Returns or sets the mode for the Hebrew spelling checker. Read/write XlHebrewModes.
- `ArabicModes As XlArabicModes  (read/write)`  
  Returns or sets the mode for the Arabic spelling checker. Read/write XlArabicModes.
- `ArabicStrictAlefHamza As Boolean  (read/write)`  
  Returns or sets whether the spelling checker uses rules regarding Arabic words beginning with an alef hamza. Read/write.
- `ArabicStrictFinalYaa As Boolean  (read/write)`  
  Returns or sets whether the spelling checker uses rules regarding Arabic words ending with the letter yaa. Read/write.
- `ArabicStrictTaaMarboota As Boolean  (read/write)`  
  Returns or sets whether the spelling checker uses rules to flag Arabic words ending with haa instead of taa marboota. Read/write.
- `RussianStrictE As Boolean  (read/write)`  
  Returns or sets whether the spelling checker uses rules regarding Russian words containing the character . Read/write.
- `SpanishModes As XlSpanishModes  (read/write)`  
  Returns or sets the mode for checking the spelling of Spanish. Read/write.
- `PortugalReform As XlPortugueseReform  (read/write)`  
  Returns or sets the mode for checking the spelling of European Portuguese. Read/write.
- `BrazilReform As XlPortugueseReform  (read/write)`  
  Returns or sets the mode for checking the spelling of Brazilian Portuguese. Read/write.
