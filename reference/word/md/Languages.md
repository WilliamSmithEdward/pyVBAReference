# Languages

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002096E-0000-0000-C000-000000000046}  

A collection of Language objects that represent languages used for proofing or formatting in Word.

**Remarks:** Use the Languages property to return the Languages collection. The following example displays the localized name for each language. Use Languages (index) to return a single Language object, where index can be the value of the Name property, the value of the NameLocal property, one of the WdLanguageID constants, or one of the MsoLanguageID constants. (For the list of valid MsoLanguageID constants, see the Object Browser in the Visual Basic Editor.) The Count property returns the number of languages for which you can mark text (languages for which proofing tools are available). To check proofing, you must install the appropriate tools for each language you intend to check. You need both a .dll file and an .lex file for each of the following: the thesaurus, spelling checker, grammar checker, and hyphenation tools. If you mark text as wdNoProofing, Word skips the marked text when running a spelling or grammar check. To mark text for a specified language or for no proofing, use the Set Language command.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of languages in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Languages object.

## Methods (1)

- `Item(Index As Variant) As Language`  
  Returns an individual Language object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
