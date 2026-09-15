# BuildingBlocks

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {C6D50987-25D7-408A-BFF2-90BF86A24E93}  

Represents a collection of BuildingBlock objects for a specific building block type and category in a template.

**Remarks:** Use the Add method to create a new building block and add it to a template. The following example adds the selected text to the watermarks building block gallery of the first template in the Templates collection. The collection returned with the BuildingBlocks collection is a filtered collection based on the type and category. Depending on how you access the collection, the collection returned changes. For example, if you access a collection of building blocks with a type of wdTypeAutoText with a category of "General", the returned collection may be different from the collection returned if you access a collection of building blocks with a type of wdTypeAutoText with a category of "Custom". It is also different from the collection returned if you access the collection of building blocks with a type of wdTypeCustomAutoText with a category of "General". For more information about building blocks, see Working with Building Blocks.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified BuildingBlocks object.
- `Count As Long  (read-only)`  
  Returns the number of items in the BuildingBlocks collection. Read-only Long.

## Methods (2)

- `Item(Index As Variant) As BuildingBlock`  
  Returns a BuildingBlock object that represents the specified item in the collection.
    - `Index As Variant` (required): Specifies a String or Integer that represents the name or ordinal position of the object in the collection.
- `Add(Name As String, Range As Range, [Description As Variant], [InsertOptions As WdDocPartInsertOptions]) As BuildingBlock`  
  Creates a new building block and returns a BuildingBlock object.
    - `Name As String` (required): Specifies the name of the building block entry. Corresponds to the Name property of the BuildingBlock object.
    - `Range As Range` (required): Specifies the value of the buildling block entry. Corresponds to the Value property of the BuildingBlock object.
    - `Description As Variant` (optional): Specifies the description of the buildling block entry. Corresponds to the Description property of the BuildingBlock object.
    - `InsertOptions As WdDocPartInsertOptions` (optional): Specifies whether the building block entry is inserted as a page, a paragraph, or inline. If omitted, the default value is wdInsertContent. Corresponds to the InsertOptions property for the BuildingBlock object.
