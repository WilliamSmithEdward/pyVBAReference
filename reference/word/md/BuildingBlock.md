# BuildingBlock

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {BFD3FC23-F763-4FF8-826E-1AFBF598A4E7}  

Represents a building block in a template. A building block is pre-built content, similar to autotext, that may contain text, images, and formatting.

**Remarks:** Each BuildingBlock object is a member of the BuildingBlocks and BuildingBlockEntries collections. Building blocks are stored in Microsoft Word templates. Therefore, to access the building blocks available for a document, you need to access an attached template. Built-in building blocks are stored in the template named "Building Blocks.dotx". Use the Item method of the collection or the BuildingBlocks collection to return an individual building block. The following example accesses the first building block in the first template in the Templates collection. To create a new building block, you can use the Add method for either the BuildingBlockEntries collection or the BuildingBlocks collection. However, the recommended way to create a new building block is by using the Add method for the BuildingBlockEntries collection. The following example adds the selected text to the watermarks building block gallery of the first template in the Templates collection. Use the Insert method to insert a new building block into a document. The following example inserts the first building block in the first template into the active document at the Insertion Point.

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified BuildingBlock object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read/write)`  
  Returns or sets a String that represents the name of a building block. Read/write.
- `Type As BuildingBlockType  (read-only)`  
  Returns a BuildingBlockType object that represents the type for a building block. Read-only.
- `Description As String  (read/write)`  
  Returns or sets a String that represents the description for a building block. Read/write.
- `ID As String  (read-only)`  
  Returns a String that represents the internal identification number for a building block. Read-only.
- `Category As Category  (read-only)`  
  Returns a Category object that represents the category for a building block. Read-only.
- `Value As String  (read/write)`  
  Returns or sets a String that represents the contents of a building block. Read/write.
- `InsertOptions As Long  (read/write)`  
  Returns or sets a Long that represents how to insert the contents of a building block into a document. Read/write.

## Methods (2)

- `Delete()`  
  Deletes the specified building block.
- `Insert(Where As Range, [RichText As Variant]) As Range`  
  Inserts the value of a building block into a document and returns a Range object that represents the contents of the building block within the document.
    - `Where As Range` (required): The location of where to place the contents of the building block.
    - `RichText As Variant` (optional): True inserts the building block as rich, formatted text. False inserts the building block as plain text.
