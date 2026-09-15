# TextInput

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020927-0000-0000-C000-000000000046}  

Represents a single text form field.

**Remarks:** Use FormFields (Index), where Index is either the bookmark name associated with the text form field or the index number, to return a FormField object. Use the TextInput property with the FormField object to return a TextInput object. The following example deletes the contents of the text form field named "Text1" in the active document. The index number represents the position of the form field in the FormFields collection. The following example checks the type of the first form field in the active document. If the form field is a text form field, the example sets "Mission Critical" as the value of the field. The following example determines whether the ffield variable represents a valid text form field in the active document before it sets the default text. Use the Add method with the FormFields object to add a text form field. The following example adds a text form field at the beginning of the active document and then sets the name of the form field to "FirstName."

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TextInput object.
- `Valid As Boolean  (read-only)`  
  True if the specified form field object is a valid check box form field. Read-only Boolean. .
- `Default As String  (read/write)`  
  Returns or sets the text that represents the default text box contents. Read/write String. The string length is limited to 255 characters.
- `Type As WdTextFormFieldType  (read-only)`  
  Returns the type of text form field. Read-only WdTextFormFieldType.
- `Format As String  (read-only)`  
  Returns the text formatting for the specified text box. Read-only String.
- `Width As Long  (read/write)`  
  Returns or sets the width, in points, of the specified text input field. Read/write Long.

## Methods (2)

- `Clear()`  
  Deletes the text from the specified text form field.
- `EditType(Type As WdTextFormFieldType, [Default As Variant], [Format As Variant], [Enabled As Variant])`  
  Sets options for the specified text form field.
    - `Type As WdTextFormFieldType` (required): The text box type.
    - `Default As Variant` (optional): The default text that appears in the text box.
    - `Format As Variant` (optional): The formatting string used to format the text, number, or date (for example, "0.00," "Title Case," or "M/d/yy"). For more examples of formats, see the list of formats for the specified text form field type in the Text Form Field Options dialog box.
    - `Enabled As Variant` (optional): True to enable the form field for text entry.
