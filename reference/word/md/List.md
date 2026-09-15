# List

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020992-0000-0000-C000-000000000046}  

Represents a single list format that's been applied to specified paragraphs in a document. The List object is a member of the Lists collection.

**Remarks:** Use Lists (Index), where Index is the index number, to return a single List object. The following example returns the number of items in list one in the active document. To return all the paragraphs that have list formatting, use the ListParagraphs property. To return them as a range, use the Range property. To apply a different list format to an existing list, use the ApplyListTemplate method with the List object. To add a new list to a document, use the ApplyListTemplate method with the ListFormat object for a specified range. Use the CanContinuePreviousList method to determine whether you can continue the list formatting from a list that was previously applied to the document. Use the CountNumberedItems method to return the number of items in a numbered or bulleted list, including LISTNUM fields. To determine whether a list contains more than one list template, use the SingleListTemplate property. You can manipulate the individual List objects within a document, but for more precise control you should work with the ListFormat object.

## Properties (7)

- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained in the specified object.
- `ListParagraphs As ListParagraphs  (read-only)`  
  Returns a ListParagraphs collection that represents all the numbered paragraphs in the list, document, or range. Read-only.
- `SingleListTemplate As Boolean  (read-only)`  
  True if the entire list uses the same list template. Read-only Boolean.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified List object.
- `StyleName As String  (read-only)`  
  Returns the name of the style applied to the specified AutoText entry. Read-only String.

## Methods (6)

- `ConvertNumbersToText([NumberType As Variant])`  
  Changes the list numbers and LISTNUM fields in the specified List object.
- `RemoveNumbers([NumberType As Variant])`  
  Removes numbers or bullets from the specified list.
    - `NumberType As Variant` (optional): The type of number to be removed.
- `CountNumberedItems([NumberType As Variant], [Level As Variant]) As Long`  
  Returns the number of bulleted or numbered items and LISTNUM fields in the specified List object.
- `CanContinuePreviousList(ListTemplate As ListTemplate) As WdContinue`  
  Returns a WdContinue constant (wdContinueDisabled, wdResetList, or wdContinueList) that indicates whether the formatting from the previous list can be continued.
    - `ListTemplate As ListTemplate` (required): A list template that's been applied to previous paragraphs in the document.
- `ApplyListTemplate(ListTemplate As ListTemplate, [ContinuePreviousList As Variant], [DefaultListBehavior As Variant])`  
  Applies a set of list-formatting characteristics to the specified ListFormat object.
    - `ListTemplate As ListTemplate` (required): The list template to be applied.
    - `ContinuePreviousList As Variant` (optional): True to continue the numbering from the previous list; False to start a new list.
    - `DefaultListBehavior As Variant` (optional): Sets a value that specifies whether Microsoft Word uses new web-oriented formatting for better list display. Can be either of the following WdDefaultListBehavior constants: wdWord8ListBehavior (use formatting compatible with Microsoft Word 97) or wdWord9ListBehavior (use web-oriented formatting). For compatibility reasons, the default constant is wdWord8ListBehavior, but in new procedures you should use wdWord9ListBehavior to take advantage of improved web-oriented formatting with respect to indenting and multilevel lists.
- `ApplyListTemplateWithLevel(ListTemplate As ListTemplate, [ContinuePreviousList As Variant], [DefaultListBehavior As Variant], [ApplyLevel As Variant])`  
  Applies a set of list-formatting characteristics, optionally for a specified level.
    - `ListTemplate As ListTemplate` (required): The list template to be applied.
    - `ContinuePreviousList As Variant` (optional): True to continue the numbering from the previous list; False to start a new list.
    - `DefaultListBehavior As Variant` (optional): Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following WdDefaultListBehavior constants: wdWord8ListBehavior (use formatting compatible with Microsoft Word 97) or wdWord9ListBehavior (use Web-oriented formatting). For compatibility reasons, the default constant is wdWord8ListBehavior, but in new procedures you should use wdWord9ListBehavior to take advantage of improved Web-oriented formatting for indenting and multiple-level lists.
    - `ApplyLevel As Variant` (optional): The level to which the list template is to be applied.
