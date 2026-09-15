# SynonymInfo

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002099B-0000-0000-C000-000000000046}  

Represents the information about synonyms, antonyms, related words, or related expressions for the specified range or a given string.

**Remarks:** Use the SynonymInfo property to return a SynonymInfo object. The SynonymInfo object can be returned either from a range or from Microsoft Office Word. If it is returned from Word, you specify the lookup word or phrase and a proofing language ID. If it is returned from a range, Word uses the specified range as the lookup word. The following example returns a SynonymInfo object from Word. The following example returns a SynonymInfo object from a range. The Found property, used in the preceding examples, returns True if any information is found in the thesaurus for the specified range or for Word. Note, however, that this property returns True not only if synonyms are found but also if related words, related expressions, or antonyms are found. Many of the properties of the SynonymInfo object return a Variant that contains an array of strings. When working with these properties, you can assign the returned array to a variable and then index the variable to see the elements in the array. In the following example, _Slist_ is assigned the synonym list for the first meaning of the selected word or phrase.

## Properties (12)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified SynonymInfo object.
- `Word As String  (read-only)`  
  Returns the word or phrase that was looked up by the thesaurus. Read-only String.
- `Found As Boolean  (read-only)`  
  True if the thesaurus finds synonyms, antonyms, related words, or related expressions for the word or phrase. Read-only Boolean.
- `MeaningCount As Long  (read-only)`  
  Returns the number of entries in the list of meanings found in the thesaurus for the word or phrase. Returns 0 (zero) if no meanings were found. Read-only Long.
- `MeaningList As Variant  (read-only)`  
  Returns the list of meanings for the word or phrase. The list is returned as an array of strings. Read-only Variant.
- `PartOfSpeechList As Variant  (read-only)`  
  Returns a list of the parts of speech corresponding to the meanings found for the word or phrase looked up in the thesaurus. The list is returned as an array of integers. Read-only Variant.
- `SynonymList As Variant  (read-only)`  
  Returns a list of synonyms for a specified meaning of a word or phrase. The list is returned as an array of strings. Read-only Variant.
- `AntonymList As Variant  (read-only)`  
  Returns a list of antonyms for the word or phrase. The list is returned as an array of strings. Read-only Variant.
- `RelatedExpressionList As Variant  (read-only)`  
  Returns a list of expressions related to the specified word or phrase. The list is returned as an array of strings. Read-only Variant.
- `RelatedWordList As Variant  (read-only)`  
  Returns a list of words related to the specified word or phrase. The list is returned as an array of strings. Read-only Variant.
