# SpellingSuggestion

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209AB-0000-0000-C000-000000000046}  

Represents a single spelling suggestion for a misspelled word. The SpellingSuggestion object is a member of the SpellingSuggestions collection. The SpellingSuggestions collection includes all the suggestions for a specified word or for the first word in the specified range.

**Remarks:** Use GetSpellingSuggestions (Index), where Index is the index number, to return a single SpellingSuggestion object. The following example checks to see whether there are any spelling suggestions for the first word in the active document. If there are, the first suggestion is displayed in a message box. The Count property for the SpellingSuggestions object returns 0 (zero) if the word is spelled correctly or if there are no suggestions.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified SpellingSuggestion object.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
