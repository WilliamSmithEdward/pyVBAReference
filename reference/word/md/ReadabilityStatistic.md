# ReadabilityStatistic

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209AF-0000-0000-C000-000000000046}  

Represents one of the readability statistics for a document or range. The ReadabilityStatistic object is a member of the ReadabilityStatistics collection.

**Remarks:** Use ReadabilityStatistics (Index), where Index is the index number, to return a single ReadabilityStatistic object. The statistics are ordered as follows: Words, Characters, Paragraphs, Sentences, Sentences per Paragraph, Words per Sentence, Characters per Word, Passive Sentences, Flesch Reading Ease, and Flesch-Kincaid Grade Level. The following example returns the character count for the active document.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ReadabilityStatistic object.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Value As Single  (read-only)`  
  Returns the value of the grammar statistic. Read-only Long.
