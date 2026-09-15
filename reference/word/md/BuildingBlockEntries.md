# BuildingBlockEntries

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {39709229-56A0-4E29-9112-B31DD067EBFD}  

Represents a collection of all BuildingBlock objects in a template.

**Remarks:** Use the Add method to create a new building block and add it to a template. The following example adds the selected text to the watermarks building block gallery of the first template in the Templates collection. Unlike the Add method for the BuildingBlocks collection, you need to specify the type and category when you add a building block using the Add method of the BuildingBlockEntries collection. This is because building blocks are organized by using types and categories. When you use the BuildingBlockEntries collection, you are accessing the entire collection of building blocks in a template; however, when you use the BuildingBlocks collection, you are accessing the collection of building blocks for a specific type and category in a template. For more information about building blocks, see Working with Building Blocks.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified BuildingBlockEntries object.
- `Count As Long  (read-only)`  
  Returns the number of items in the BuildingBlockEntries collection. Read-only Long.

## Methods (2)

- `Item(Index As Variant) As BuildingBlock`  
  Returns a BuildingBlock object that represents the specified item in the collection.
    - `Index As Variant` (required): Specifies a String or Integer that represents the name or ordinal position of the object in the collection.
- `Add(Name As String, Type As WdBuildingBlockTypes, Category As String, Range As Range, [Description As Variant], [InsertOptions As WdDocPartInsertOptions]) As BuildingBlock`  
  Creates a new building block entry in a template and returns a BuildingBlock object that represents the new building block entry.
    - `Name As String` (required): Specifies the name of the building block entry. Corresponds to the Name property of the BuildingBlock object.
    - `Type As WdBuildingBlockTypes` (required): Specifies the type of building block to create. Corresponds to the Type property of the BuildingBlock object.
    - `Category As String` (required): Specifies the category of the new building block entry. Corresponds to the Category property of the BuildingBlock object.
    - `Range As Range` (required): Specifies the value of the buildling block entry. Corresponds to the Value property of the BuildingBlock object.
    - `Description As Variant` (optional): Specifies the description of the buildling block entry. Corresponds to the Description property of the BuildingBlock object.
    - `InsertOptions As WdDocPartInsertOptions` (optional): Specifies whether the building block entry is inserted as a page, a paragraph, or inline. If omitted, the default value is wdInsertContent. Corresponds to the InsertOptions property for the BuildingBlock object.
