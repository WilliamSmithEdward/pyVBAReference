# DropDown

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020925-0000-0000-C000-000000000046}  

Represents a drop-down form field that contains a list of items in a form.

**Remarks:** Use FormFields (_index_), where _index_ is the index number or the bookmark name associated with the drop-down form field, to return a single FormField object. Use the DropDown property with the FormField object to return a DropDown object. The following example selects the first item in the drop-down form field named "DropDown" in the active document. The index number represents the position of the form field in the FormFields collection. The following example checks the type of the first form field in the active document. If it is a drop-down form field, the second item is selected. The following example determines whether form field represented by _ffield_ is a valid drop-down form field before adding an item to it. Use the Add method with the FormFields collection to add a drop-down form field. The following example adds a drop-down form field at the beginning of the active document and then adds items to the form field.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified DropDown object.
- `Valid As Boolean  (read-only)`  
  True if the specified form field object is a valid drop down form field. Read-only Boolean.
- `Default As Long  (read/write)`  
  Returns or sets a Long that represents the default drop-down item. Read/write.
- `Value As Long  (read/write)`  
  Returns or sets the number of the selected item in a drop-down form field. Read/write Long.
- `ListEntries As ListEntries  (read-only)`  
  Returns a ListEntries collection that represents all the items in a DropDown object.
