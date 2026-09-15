# Hyperlinks

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002099C-0000-0000-C000-000000000046}  

Represents the collection of Hyperlink objects in a document, range, or selection.

**Remarks:** Use the Hyperlinks property to return the Hyperlinks collection. The following example checks all the hyperlinks in document one for a link that contains the word "Microsoft" in the address. If a hyperlink is found, it is activated with the Follow method. Use the Add method to create a hyperlink and add it to the Hyperlinks collection. The following example creates a new hyperlink to the MSN Web site. Use Hyperlinks (Index), where Index is the index number, to return a single Hyperlink object in a document, range, or selection. The following example activates the first hyperlink in the selection. The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Hyperlinks object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of hyperlinks in the collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Item(Index As Variant) As Hyperlink`  
  Returns an individual Hyperlink object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Anchor As Object, [Address As Variant], [SubAddress As Variant], [ScreenTip As Variant], [TextToDisplay As Variant], [Target As Variant]) As Hyperlink`  
  Returns a Hyperlink object that represents a new hyperlink added to a range, selection, or document.
    - `Anchor As Object` (required): The text or graphic that you want turned into a hyperlink.
    - `Address As Variant` (optional): The address for the specified link. The address can be an email address, an Internet address, or a file name. Note that Microsoft Word doesn't check the accuracy of the address.
    - `SubAddress As Variant` (optional): The name of a location within the destination file, such as a bookmark, named range, or slide number.
    - `ScreenTip As Variant` (optional): The text that appears as a ScreenTip when the mouse pointer is positioned over the specified hyperlink. The default value is "Address".
    - `TextToDisplay As Variant` (optional): The display text of the specified hyperlink. The value of this argument replaces the text or graphic specified by Anchor.
    - `Target As Variant` (optional): The name of the frame or window in which you want to load the specified hyperlink.
