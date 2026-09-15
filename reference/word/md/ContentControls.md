# ContentControls

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {804CD967-F83B-432D-9446-C61A45CFEFF0}  

A collection of ContentControl objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain content such as dates, lists, or paragraphs of formatted text.

**Remarks:** Use the Add method to create a new content control and insert it into a document. The following example creates a new drop-down list content control and adds several items to the list. Use the Item method to access a specific content control in the collection. The following example accesses the third content control in the active document, and if the control is a drop-down list or a combo box, moves the first item to the bottom of the list and the last item to the top of the list. Use the ContentControl object to work with individual content controls. For more information, see Working with Content Controls.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ContentControls object.
- `Count As Long  (read-only)`  
  Returns the number of items in the ContentControls collection. Read-only Long.

## Methods (2)

- `Item(Index As Variant) As ContentControl`  
  Returns a ContentControl object that represents the specified content control within the collection of content controls in a document.
    - `Index As Variant` (required): Specifies the ordinal position of the content control to return.
- `Add([Type As WdContentControlType], [Range As Variant]) As ContentControl`  
  Adds a new content control, of the type specified, into the active document and returns a ContentControl object that represents the new content control.
    - `Type As WdContentControlType` (optional): Specifies the type of content control to insert into the active document. If omitted, Microsoft Word inserts a rich-text content control.
    - `Range As Variant` (optional): Specifies where in the active document to place the content control. If omitted, Word places the content control at the position of the insertion point or replaces the current selection.
