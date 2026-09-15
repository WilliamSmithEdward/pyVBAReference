# FormField

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020928-0000-0000-C000-000000000046}  

Represents a single form field. The FormField object is a member of the FormFields collection.

**Remarks:** Use FormFields (_index_), where _index_ is a bookmark name or index number, to return a single FormField object. The following example sets the result of the Text1 form field to "Don Funk." The index number represents the position of the form field in the selection, range, or document. The following example displays the name of the first form field in the selection. Use the Add method with the FormFields object to add a form field. The following example adds a check box at the beginning of the active document and then selects the check box. Use the CheckBox, DropDown, and TextInput properties with the FormField object to return the CheckDown, DropDown, and TextInput objects. The following example selects the check box named "Check1."

## Properties (20)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FormField object.
- `Type As WdFieldType  (read-only)`  
  Returns the field type. Read-only WdFieldType.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write String.
- `EntryMacro As String  (read/write)`  
  Returns or sets an entry macro name for the specified form field (CheckBox, DropDown, or TextInput). Read/write String.
- `ExitMacro As String  (read/write)`  
  Returns or sets an exit macro name for the specified form field (CheckBox, DropDown, or TextInput). Read/write String.
- `OwnHelp As Boolean  (read/write)`  
  Specifies the source of the text that's displayed in a message box when a form field has the focus and the user presses F1. Read/write Boolean.
- `OwnStatus As Boolean  (read/write)`  
  Specifies the source of the text that's displayed in the status bar when a form field has the focus. Read/write Boolean.
- `HelpText As String  (read/write)`  
  Returns or sets the text that's displayed in a message box when the form field has the focus and the user presses F1. Read/write String.
- `StatusText As String  (read/write)`  
  Returns or sets the text that is displayed in the status bar when a form field has the focus. Read/write String.
- `Enabled As Boolean  (read/write)`  
  True if a form field is enabled. Read/write Boolean.
- `Result As String  (read/write)`  
  Returns a String that represents the result of the specified form field. Read/write.
- `TextInput As TextInput  (read-only)`  
  Returns a TextInput object that represents a text form field.
- `CheckBox As CheckBox  (read-only)`  
  Returns a CheckBox object that represents a check box form field. Read-only.
- `DropDown As DropDown  (read-only)`  
  Returns a DropDown object that represents a drop-down form field. Read-only.
- `Next As FormField  (read-only)`  
  Returns the next form field in the collection. Read-only.
- `Previous As FormField  (read-only)`  
  Returns the previous form field in the collection. Read-only.
- `CalculateOnExit As Boolean  (read/write)`  
  True if references to the specified form field are automatically updated whenever the field is exited. Read/write Boolean.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within the form field.

## Methods (4)

- `Select()`  
  Selects the specified object.
- `Copy()`  
  Copies the specified form field to the Clipboard.
- `Cut()`  
  Removes the specified form field from the document and places it on the Clipboard.
- `Delete()`  
  Deletes the specified form field.
