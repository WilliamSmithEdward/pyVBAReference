# Bibliography

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {3834F60F-EE8C-455D-A441-D766675D6D3B}  

Represents the list of sources available attached to the document (in the current list) or the list of sources available in the application (in the master list).

**Remarks:** Biblographies are composed of sources. Also see the Source object and the Sources collection. For more information, see Working with Bibliographies.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Bibliography object.
- `Sources As Sources  (read-only)`  
  Returns a Sources collection that represents all the sources contained in a bibliography. Read-only.
- `BibliographyStyle As String  (read/write)`  
  Returns or sets a String that represents the name of the active style to use for a bibliography. Read/write.

## Methods (1)

- `GenerateUniqueTag() As String`  
  Generates a unique identification tag for a bibliography source and returns a String that represents the tag.
