# Source

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {4A6AE865-199D-4EA3-9F6B-125BD9C40EDF}  

Represents an individual source, such as a book, journal article, or interview.

**Remarks:** For more information, see Working with Bibliographies.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Source object.
- `Tag As String  (read-only)`  
  Returns a String that represents an internal identifying label for a source. Read-only.
- `Field As String  (read/write)`  
  Returns a String that represents the value of a field in a bibliography source. Read-only.
- `XML As String  (read-only)`  
  Returns a String that represents the XML markup for a Source object. Read-only.
- `Cited As Boolean  (read-only)`  
  Returns a Boolean that represents whether a source has been cited in a document. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the specified source.
