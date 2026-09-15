# ListFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209C0-0000-0000-C000-000000000046}  

Represents the list formatting attributes that can be applied to the paragraphs in a range.

**Remarks:** Use the ListFormat property to return the ListFormat object for a range. The following example applies the default bulleted list format to the selection. An easy way to apply list formatting is to use the ApplyBulletDefault, ApplyNumberDefault, and ApplyOutlineNumberDefault methods, which correspond, respectively, to the first list format (excluding None) on each tab in the Bullets and Numbering dialog box. To apply a format other than the default format, use the ApplyListTemplate method, which allows you to specify the list format (list template) you want to apply. Use the List or ListTemplate property to return the list or list template from the first paragraph in the specified range. Use the ListFormat property with a Range object to access the list formatting properties and methods available for the specified range. The following example applies the default bulleted list format to the second paragraph in the active document. However, if there is already a list defined in your document, you can access a List object by using the Lists property.

## Properties (12)

- `ListLevelNumber As Long  (read/write)`  
  Returns or sets the list level for the first paragraph in the specified ListFormat object. Read/write Long.
- `List As List  (read-only)`  
  Returns a List object that represents the first formatted list contained in the specified ListFormat object.
- `ListTemplate As ListTemplate  (read-only)`  
  Returns a ListTemplate object that represents the list formatting for the specified ListFormat object.
- `ListValue As Long  (read-only)`  
  Returns the numeric value of the first paragraph in the range for the specified ListFormat object. Read-only Long.
- `SingleList As Boolean  (read-only)`  
  True if the specified ListFormat object contains only one list. Read-only Boolean.
- `SingleListTemplate As Boolean  (read-only)`  
  True if the entire ListFormat object uses the same list template. Read-only Boolean.
- `ListType As WdListType  (read-only)`  
  Returns the type of lists that are contained in the range for the specified ListFormat object. Read-only WdListType.
- `ListString As String  (read-only)`  
  Returns a String that represents the appearance of the list value of the first paragraph in the range for the specified ListFormat object. For example, the second paragraph in an alphabetical list would return B. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ListFormat object.
- `ListPictureBullet As InlineShape  (read-only)`  
  Returns the InlineShape object that represents the picture used as a bullet in a picture bulleted list.

## Methods (11)

- `CanContinuePreviousList(ListTemplate As ListTemplate) As WdContinue`  
  Returns a WdContinue constant (wdContinueDisabled, wdResetList, or wdContinueList) that indicates whether the formatting from the previous list can be continued.
    - `ListTemplate As ListTemplate` (required): A list template that's been applied to previous paragraphs in the document.
- `RemoveNumbers([NumberType As Variant])`  
  Removes numbers or bullets from the specified list.
    - `NumberType As Variant` (optional): The type of number to be removed.
- `ConvertNumbersToText([NumberType As Variant])`  
  Changes the list numbers and LISTNUM fields in the specified ListFormat object to text.
- `CountNumberedItems([NumberType As Variant], [Level As Variant]) As Long`  
  Returns the number of bulleted or numbered items and LISTNUM fields in the specified ListFormat object.
- `ListOutdent()`  
  Decreases the list level of the paragraphs in the range for the specified ListFormat object, in increments of one level.
- `ListIndent()`  
  Increases the list level of the paragraphs in the range for the specified ListFormat object, in increments of one level.
- `ApplyBulletDefault([DefaultListBehavior As Variant])`  
  Adds bullets and formatting to the paragraphs in the range for the specified ListFormat object.
    - `DefaultListBehavior As Variant` (optional): Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following constants: wdWord8ListBehavior (use formatting compatible with Microsoft Word 97) or wdWord9ListBehavior (use Web-oriented formatting). For compatibility reasons, the default constant is wdWord8ListBehavior, but in new procedures you should use wdWord9ListBehavior to take advantage of improved Web-oriented formatting with respect to indenting and multilevel lists.
- `ApplyNumberDefault([DefaultListBehavior As Variant])`  
  Adds the default numbering scheme to the paragraphs in the range for the specified ListFormat object.
    - `DefaultListBehavior As Variant` (optional): Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following constants: wdWord8ListBehavior (use formatting compatible with Microsoft Word 97) or wdWord9ListBehavior (use Web-oriented formatting). For compatibility reasons, the default constant is wdWord8ListBehavior, but in new procedures you should use wdWord9ListBehavior to take advantage of improved Web-oriented formatting with respect to indenting and multilevel lists.
- `ApplyOutlineNumberDefault([DefaultListBehavior As Variant])`  
  Adds the default outline-numbering scheme to the paragraphs in the range for the specified ListFormat object.
    - `DefaultListBehavior As Variant` (optional): Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following constants: wdWord8ListBehavior (use formatting compatible with Microsoft Word 97) or wdWord9ListBehavior (use Web-oriented formatting). For compatibility reasons, the default constant is wdWord8ListBehavior, but in new procedures you should use wdWord9ListBehavior to take advantage of improved Web-oriented formatting with respect to indenting and multilevel lists.
- `ApplyListTemplate(ListTemplate As ListTemplate, [ContinuePreviousList As Variant], [ApplyTo As Variant], [DefaultListBehavior As Variant])`  
  Applies a set of list-formatting characteristics to the specified ListFormat object.
    - `ListTemplate As ListTemplate` (required): The list template to be applied.
    - `ContinuePreviousList As Variant` (optional): True to continue the numbering from the previous list; False to start a new list.
    - `ApplyTo As Variant` (optional): The portion of the list that the list template is to be applied to. Can be one of the following WdListApplyTo constants: wdListSelection, wdListWholeList, or wdListThisPointForward.
    - `DefaultListBehavior As Variant` (optional): Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following WdDefaultListBehavior constants: wdWord8ListBehavior (use formatting compatible with Microsoft Word 97) or wdWord9ListBehavior (use Web-oriented formatting). For compatibility reasons, the default constant is wdWord8ListBehavior, but in new procedures you should use wdWord9ListBehavior to take advantage of improved Web-oriented formatting with respect to indenting and multilevel lists.
- `ApplyListTemplateWithLevel(ListTemplate As ListTemplate, [ContinuePreviousList As Variant], [ApplyTo As Variant], [DefaultListBehavior As Variant], [ApplyLevel As Variant])`  
  Applies a set of list-formatting characteristics, optionally for a specified level.
    - `ListTemplate As ListTemplate` (required): The list template to be applied.
    - `ContinuePreviousList As Variant` (optional): True to continue the numbering from the previous list; False to start a new list.
    - `ApplyTo As Variant` (optional): One of the WdListApplyTo constants that specifies the portion of the list that the list template will be applied to.
    - `DefaultListBehavior As Variant` (optional): Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following WdDefaultListBehavior constants: wdWord8ListBehavior (use formatting compatible with Microsoft Word 97) or wdWord9ListBehavior (use Web-oriented formatting). For compatibility reasons, the default constant is wdWord8ListBehavior, but in new procedures you should use wdWord9ListBehavior to take advantage of improved Web-oriented formatting for indenting and multiple-level lists.
    - `ApplyLevel As Variant` (optional): The level to which the list template is to be applied.
