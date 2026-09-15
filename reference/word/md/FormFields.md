# FormFields

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020929-0000-0000-C000-000000000046}  

A collection of FormField objects that represent all the form fields in a selection, range, or document.

**Remarks:** Use the FormFields property to return the FormFields collection. The following example counts the number of text box form fields in the active document. Use the Add method with the FormFields object to add a form field. The following example adds a check box at the beginning of the active document and then selects the check box. Use FormFields (Index), where Index is a bookmark name or index number, to return a single FormField object. The following example sets the result of the Text1 form field to "Don Funk." The index number represents the position of the form field in the selection, range, or document. The following example displays the name of the first form field in the selection.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FormFields object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of fields in the collection. Read-only.
- `Shaded As Boolean  (read/write)`  
  True if shading is applied to form fields. Read/write Boolean.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Item(Index As Variant) As FormField`  
  Returns an individual FormField object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Range As Range, Type As WdFieldType) As FormField`  
  Returns a FormField object that represents a new form field added at a range.
    - `Range As Range` (required): The range where you want to add the form field. If the range isn't collapsed, the form field replaces the range.
    - `Type As WdFieldType` (required): The type of form field to add.
