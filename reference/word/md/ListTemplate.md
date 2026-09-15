# ListTemplate

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002098F-0000-0000-C000-000000000046}  

Represents a single list template that includes all the formatting that defines a list. The ListTemplate object is a member of the ListTemplates collection.

**Remarks:** Each of the seven formats (excluding None) found on each of the three tabs in the Bullets and Numbering dialog box corresponds to a list template object. These predefined list templates can be accessed from the three ListGallery objects in the ListGalleries collection. Documents and templates can also contain collections of list templates. Use ListTemplates (Index), where Index is a number from 1 through 7, to return a single list template from a list gallery. The following example returns the third list format (excluding None) on the Numbered tab in the Bullets and Numbering dialog box. The following example sets an object variable equal to the list template used in the third list in the active document, and then it applies that list template to the selection. Use the Add method to add a list template to the collection of list templates in a document or template. To see whether the specified list template contains the formatting built into Word, use the Modified property with the ListGallery object. To reset formatting to the original list format, use the Reset method for the ListGallery object.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ListTemplate object.
- `OutlineNumbered As Boolean  (read/write)`  
  True if the specified ListTemplate object is outline numbered. Read/write Boolean.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write String.
- `ListLevels As ListLevels  (read-only)`  
  Returns a ListLevels collection that represents all the levels for the specified ListTemplate.

## Methods (1)

- `Convert([Level As Variant]) As ListTemplate`  
  Converts a multiple-level list to a single-level list, or vice versa.
    - `Level As Variant` (optional): The level to use for formatting the new list. When converting a multiple-level list to a single-level list, this argument can be a number from 1 through 9. When converting a single-level list to a multiple-level list, 1 is the only valid value. If this argument is omitted, 1 is the default value.
