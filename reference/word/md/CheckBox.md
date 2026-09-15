# CheckBox

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020926-0000-0000-C000-000000000046}  

Represents a single check box form field.

**Remarks:** Use FormFields (Index), where Index is index number or the bookmark name associated with the check box, to return a single FormField object. Use the CheckBox property with the FormField object to return a CheckBox object. The following example selects the check box form field named "Check1" in the active document. The index number represents the position of the form field in the FormFields collection. The following example checks the type of the first form field; if it is a check box, the check box is selected. The following example determines whether the ffield object is valid before changing the check box size to 14 points. Use the Add method with the FormFields object to add a check box form field. The following example adds a check box at the beginning of the active document, sets the name to "Color", and then selects the check box.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CheckBox object.
- `Valid As Boolean  (read-only)`  
  True if the specified form field object is a valid check box form field. Read-only Boolean.
- `AutoSize As Boolean  (read/write)`  
  True sizes the check box or text frame according to the font size of the surrounding text. False sizes the check box or text frame according to the Size property. Read/write Boolean.
- `Size As Single  (read/write)`  
  Returns or sets the size of a check box, in points. Read/write Single.
- `Default As Boolean  (read/write)`  
  Returns or sets the default check box value. True if the default value is checked. Read/write Boolean.
- `Value As Boolean  (read/write)`  
  True if the check box is selected. Read/write Boolean.
