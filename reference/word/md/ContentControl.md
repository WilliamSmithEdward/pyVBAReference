# ContentControl

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {EE95AFE3-3026-4172-B078-0E79DAB5CC3D}  

An individual content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as dates, lists, or paragraphs of formatted text. The ContentControl object is a member of the ContentControls collection.

**Remarks:** Use the Add method of the ContentControls collection to create a content control. Use the Type parameter of the Add method to specify the type of content control to create. The following example create a new drop-down list content control and adds several items to the list. Use the Type property to change the content control to a different type of content control. For example, perhaps you want to change from a date control to a text control. However, you may not be able to change all content controls to another type; some may not allow changing their type. In addition, depending on the contents of a content control, you may not be able to change the type. For example, if the content control that you want to change to does not allow the type of content that is in the existing content control, attempting to change the type is not allowed and generates a run-time error. The following example inserts a date content control and sets the value of the control, and then changes the control to a text content control. Use the SetPlaceholderText method to change the placeholder text from the default string to something more appropriate for the control.

## Properties (31)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ContentControl object.
- `Range As Range  (read-only)`  
  Returns a Range that represents the contents of the content control in the active document. Read-only.
- `LockContentControl As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the user can delete a content control from the active document. Read/write.
- `LockContents As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the user can edit the contents of a content control. Read/write.
- `XMLMapping As XMLMapping  (read-only)`  
  Returns an XMLMapping object that represents the mapping of a content control to XML data in the data store of a document. Read-only.
- `Type As WdContentControlType  (read/write)`  
  Returns or sets a WdContentControlType that represents the type for a content control. Read/write.
- `DropdownListEntries As ContentControlListEntries  (read-only)`  
  Returns a ContentControlListEntries collection that represents the items in a drop-down list content control or in a combo box content control. Read-only.
- `PlaceholderText As BuildingBlock  (read-only)`  
  Returns a BuildingBlock object that represents the placeholder text for a content control. Read-only.
- `Title As String  (read/write)`  
  Returns or sets a String that represents the title for a content control. Read/write.
- `DateDisplayFormat As String  (read/write)`  
  Returns or sets a String that represents the format in which dates are displayed. Read/write.
- `MultiLine As Boolean  (read/write)`  
  Returns a Boolean that represents whether a text content control allows multiple lines of text. Read/write.
- `ParentContentControl As ContentControl  (read-only)`  
  Returns a ContentControl that represents the parent content control for a content control that is nested inside a rich-text control or group control. Read-only.
- `Temporary As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to remove a content control from the active document when the user edits the contents of the control. Read/write.
- `ID As String  (read-only)`  
  Returns a String that represents the identification for a content control. Read-only.
- `ShowingPlaceholderText As Boolean  (read-only)`  
  Returns a Boolean that indicates whether the placeholder text for the content control is displayed. Read-only.
- `DateStorageFormat As WdContentControlDateStorageFormat  (read/write)`  
  Returns or sets a WdContentControlDateStorageFormat that represents the format for storage and retrieval of dates when a date content control is bound to the XML data store of the active document. Read/write.
- `BuildingBlockType As WdBuildingBlockTypes  (read/write)`  
  Returns or sets a WdBuildingBlockTypes constant that represents they type of building block for a building block content control. Read/write.
- `BuildingBlockCategory As String  (read/write)`  
  Returns or sets a String that represents the category for a building block content control. Read/write.
- `DateDisplayLocale As WdLanguageID  (read/write)`  
  Returns a WdLanguageID that represents the language format for the date displayed in a date content control. Read/write.
- `DefaultTextStyle As Variant  (read/write)`  
  Returns or sets a Variant that represents the name of the character style to use to format text in a text content control. Read/write.
- `DateCalendarType As WdCalendarType  (read/write)`  
  Returns or sets a WdCalendarType constant that represents the calendar type for a calendar content control. Read/write.
- `Tag As String  (read/write)`  
  Returns or sets a String that represents a value to identify a content control. Read/write.
- `Checked As Boolean  (read/write)`  
  Returns or sets a Boolean that represents the current state (checked/unchecked) for a check box. Read/write.
- `Color As WdColor  (read/write)`  
  Returns or sets the color of the content control. Read/write WdColor.
- `Appearance As WdContentControlAppearance  (read/write)`  
  Returns or sets the appearance of the content control. Read/write WdContentControlAppearance.
- `Level As WdContentControlLevel  (read-only)`  
  Returns the level of the content control-whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline. Read-only WdContentControlLevel.
- `RepeatingSectionItems As RepeatingSectionItemColl  (read-only)`  
  Returns the collection of repeating section items in the specified repeating section content control. Read-only.
- `RepeatingSectionItemTitle As String  (read/write)`  
  Returns or sets the name of the repeating section items used in the context menu associated with the specified repeating section content control. Read/write.
- `AllowInsertDeleteSection As Boolean  (read/write)`  
  Gets or sets whether users can add or remove sections from the specified repeating section content control by using the user interface.

## Methods (7)

- `Copy()`  
  Copies the content control from the active document to the Clipboard.
- `Cut()`  
  Removes the content control from the active document and moves the content control to the Clipboard.
- `Delete([DeleteContents As Boolean])`  
  Deletes the specified content control and the contents of the content control.
    - `DeleteContents As Boolean` (optional): Specifies whether to delete the contents of the content control. True removes both the content control and its contents. False removes the control but leaves the contents of the content control in the active document. The default value is False.
- `SetPlaceholderText([BuildingBlock As BuildingBlock], [Range As Range], [Text As String])`  
  Sets the placeholder text that displays in the content control until a user enters their own text.
    - `BuildingBlock As BuildingBlock` (optional): Specifies a BuildingBlock object that contains the contents of the placeholder text.
    - `Range As Range` (optional): Specifies a Range object that contains the contents of the placeholder text.
    - `Text As String` (optional): Specifies the contents of the placeholder text.
- `Ungroup()`  
  Removes a group content control from a document so that its child content controls are no longer nested and can be freely edited.
- `SetCheckedSymbol(CharacterNumber As Long, [Font As String])`  
  Sets the symbol used to represent the checked state of a check box content control.
    - `CharacterNumber As Long` (required): The Unicode character number for the specified symbol. This value will always be the sum of 31 (the number of control symbols at the beginning of the font) and the number that corresponds to the position of the symbol in the table of symbols (counting from left to right). For example, to specify a delta character at position 37 in the table of symbols in the Symbol font, set CharacterNumber to 68.
    - `Font As String` (optional): The name of the font that contains the symbol.
- `SetUncheckedSymbol(CharacterNumber As Long, [Font As String])`  
  Sets the symbol used to represent the unchecked state of a check box content control.
    - `CharacterNumber As Long` (required): The Unicode character number for the specified symbol. This value will always be the sum of 31 (the number of control symbols at the beginning of the font) and the number that corresponds to the position of the symbol in the table of symbols (counting from left to right). For example, to specify a delta character at position 37 in the table of symbols in the Symbol font, set CharacterNumber to 68.
    - `Font As String` (optional): The name of the font that contains the symbol.
