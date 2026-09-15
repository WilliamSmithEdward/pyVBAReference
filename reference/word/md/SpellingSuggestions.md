# SpellingSuggestions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209AA-0000-0000-C000-000000000046}  

A collection of SpellingSuggestion objects that represent all the suggestions for a specified word or for the first word in the specified range.

**Remarks:** Use the GetSpellingSuggestions method to return the SpellingSuggestions collection. The SpellingSuggestions method, when applied to the Application object, must specify the word to be checked. When the GetSpellingSuggestions method is applied to a range, the first word in the range is checked. The following example checks to see whether there are any spelling suggestions for any of the words in the active document. If there are, the suggestions are displayed in message boxes. You cannot add suggestions to or remove suggestions from the collection of spelling suggestions. Spelling suggestions are derived from main and custom dictionary files.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified SpellingSuggestions object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of spelling suggestions in the collection. Read-only.
- `SpellingErrorType As WdSpellingErrorType  (read-only)`  
  Returns the spelling error type. Read-only WdSpellingErrorType.

## Methods (1)

- `Item(Index As Long) As SpellingSuggestion`  
  Returns an individual SpellingSuggestion object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
