# Language

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002096D-0000-0000-C000-000000000046}  

Represents a language used for proofing or formatting in Microsoft Word. The Language object is a member of the Languages collection.

**Remarks:** Use Languages (Index) to return a single Language object, where Index can be the value of the Name property, the value of the NameLocal property, one of the WdLanguageID constants, or one of the MsoLanguageID constants. (For the list of valid WdLanguageID or MsoLanguageID constants, see the Object Browser in the Visual Basic Editor.) The Name property returns the name of a language, whereas the NameLocal property returns the name of a language in the language of the user. The following example returns the string "Italiano" for Name and "Italian (Standard)" for NameLocal when it is run in the U.S. English version of Word. For each language for which proofing tools are installed, you can use the ActiveGrammarDictionary, ActiveHyphenationDictionary, ActiveSpellingDictionary, and ActiveThesaurusDictionary properties to return the corresponding Dictionary object. The following example returns the full path for the active spelling dictionary used in the U.S. English version of Word. The writing style is the set of rules used by the grammar checker. The WritingStyleList property returns an array of strings that represent the available writing styles for the specified language.

## Properties (13)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Language object.
- `ID As WdLanguageID  (read-only)`  
  Returns a number that identifies the specified language. Read-only WdLanguageID.
- `NameLocal As String  (read-only)`  
  Returns the name of a proofing tool language in the language of the user. Read-only String.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `ActiveGrammarDictionary As Dictionary  (read-only)`  
  Returns a Dictionary object that represents the active grammar dictionary for the specified language. Read-only.
- `ActiveHyphenationDictionary As Dictionary  (read-only)`  
  Returns a Dictionary object that represents the active hyphenation dictionary for the specified language. Read-only.
- `ActiveSpellingDictionary As Dictionary  (read-only)`  
  Returns a Dictionary object that represents the active spelling dictionary for the specified language.
- `ActiveThesaurusDictionary As Dictionary  (read-only)`  
  Returns a Dictionary object that represents the active thesaurus dictionary for the specified language.
- `DefaultWritingStyle As String  (read/write)`  
  Returns or sets the default writing style used by the grammar checker for the specified language. Read/write String.
- `WritingStyleList As Variant  (read-only)`  
  Returns a string array that contains the names of all writing styles available for the specified language. Read-only Variant.
- `SpellingDictionaryType As WdDictionaryType  (read/write)`  
  Returns or sets the proofing tool type. Read/write WdDictionaryType.
