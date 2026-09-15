# Template

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002096A-0000-0000-C000-000000000046}  

Represents a document template. The Template object is a member of the Templates collection. The Templates collection includes all the available Template objects.

**Remarks:** Use Templates (Index), where Index is the template name or the index number, to return a single Template object. The following example saves the Memo2.dot template if it is in the Templates collection. The index number represents the position of the template in the Templates collection. The following example opens the first template in the Templates collection. The Add method is not available for the Templates collection. Instead, you can add a template to the Templates collection by doing any of the following: - Using the Open method with the Documents collection to open a document based on a template or a template - Using the Add method with the Documents collection to open a new document based on a template - Using the Add method with the Addins collection to load a global template - Using the AttachedTemplate property with the Document object to attach a template to a document Use the NormalTemplate property to return a template object that refers to the Normal template. Use the AttachedTemplate property to return the template attached to the specified document.

## Properties (24)

- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Template object.
- `Path As String  (read-only)`  
  Returns the path to the specified document template. Read-only String.
- `LanguageID As WdLanguageID  (read/write)`  
  Returns or sets a WdLanguageID constant that represents the language for the specified range. Read/write.
- `Saved As Boolean  (read/write)`  
  True if the specified template has not changed since it was last saved. False if Microsoft Word displays a prompt to save changes when the document is closed. Read/write Boolean.
- `Type As WdTemplateType  (read-only)`  
  Returns the template type. Read-only WdTemplateType.
- `FullName As String  (read-only)`  
  Specifies the name of a template, including the drive or Web path. Read-only String.
- `BuiltInDocumentProperties As Object  (read-only)`  
  Returns a DocumentProperties collection that represents all the built-in document properties for the specified document.
- `CustomDocumentProperties As Object  (read-only)`  
  Returns a DocumentProperties collection that represents all the custom document properties for the specified document.
- `ListTemplates As ListTemplates  (read-only)`  
  Returns a ListTemplates collection that represents all the list formats for the specified template. Read-only.
- `LanguageIDFarEast As WdLanguageID  (read/write)`  
  Returns or sets an East Asian language for the specified object. Read/write WdLanguageID.
- `VBProject As VBProject  (read-only)`  
  Returns the VBProject object for the specified template.
- `KerningByAlgorithm As Boolean  (read/write)`  
  True if Microsoft Word kerns half-width Latin characters and punctuation marks in the specified document. Read/write Boolean.
- `JustificationMode As WdJustificationMode  (read/write)`  
  Returns or sets the character spacing adjustment for the specified template. Read/write WdJustificationMode.
- `FarEastLineBreakLevel As WdFarEastLineBreakLevel  (read/write)`  
  Returns or sets the line break control level for the specified document. Read/write WdFarEastLineBreakLevel.
- `NoLineBreakBefore As String  (read/write)`  
  Returns or sets the kinsoku characters before which Microsoft Word will not break a line. Read/write String.
- `NoLineBreakAfter As String  (read/write)`  
  Returns or sets the kinsoku characters after which Microsoft Word will not break a line. Read/write String.
- `NoProofing As Long  (read/write)`  
  True if the spelling and grammar checker ignores documents based on this template. Read/write Long.
- `FarEastLineBreakLanguage As WdFarEastLineBreakLanguageID  (read/write)`  
  Returns or sets the East Asian language to use when breaking lines of text in the specified document or template. Read/write WdFarEastLineBreakLanguageID.
- `BuildingBlockEntries As BuildingBlockEntries  (read-only)`  
  Returns a BuildingBlockEntries collection that represents the collection of building block entries in a template. Read-only.
- `BuildingBlockTypes As BuildingBlockTypes  (read-only)`  
  Returns a BuildingBlockTypes collection that represents the collection of building block types that are contained in a template. Read-only.
- `AutoSaveOn As Boolean  (read/write)`

## Methods (2)

- `OpenAsDocument() As Document`  
  Opens the specified template as a document and returns a Document object.
- `Save()`  
  Saves the specified template.
