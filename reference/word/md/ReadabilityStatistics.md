# ReadabilityStatistics

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209AE-0000-0000-C000-000000000046}  

A collection of ReadabilityStatistic objects for a document or range.

**Remarks:** Use the ReadabilityStatistics property to return the ReadabilityStatistics collection. The following example enumerates the readability statistics for the selection and displays each one in a message box. Use ReadabilityStatistics (Index), where Index is the index number, to return a single ReadabilityStatistic object. The statistics are ordered as follows: Words, Characters, Paragraphs, Sentences, Sentences per Paragraph, Words per Sentence, Characters per Word, Passive Sentences, Flesch Reading Ease, and Flesch-Kincaid Grade Level. The following example returns the word count for the active document.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ReadabilityStatistics object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of readability statistics in the collection. Read-only.

## Methods (1)

- `Item(Index As Variant) As ReadabilityStatistic`  
  Returns an individual ReadabilityStatistic object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
