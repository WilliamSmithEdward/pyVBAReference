# TabStops

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020955-0000-0000-C000-000000000046}  

A collection of TabStop objects that represent the custom and default tabs for a paragraph or group of paragraphs.

**Remarks:** Use the TabStops property to return the TabStops collection. The following example clears all the custom tab stops from the first paragraph in the active document. The following example adds a tab stop positioned at 2.5 inches to the selected paragraphs and then displays the position of each item in the TabStops collection. Use the Add method to add a tab stop. The following example adds two tab stops to the selected paragraphs. The first tab stop is a left-aligned tab with a dotted tab leader positioned at 1 inch (72 points). The second tab stop is centered and is positioned at 2 inches. You can also add a tab stop by specifying a location with the TabStops property. The following example adds a right-aligned tab stop positioned at 2 inches to the selected paragraphs. Use TabStops (Index), where Index is the location of the tab stop (in points) or the index number, to return a single TabStop object. Tab stops are indexed numerically from left to right along the ruler. The following example removes the first custom tab stop from the first paragraph in the active document. The following example adds a right-aligned tab stop positioned at 2 inches to the selected paragraphs.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of tab stops in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TabStops object.

## Methods (5)

- `Item(Index As Variant) As TabStop`  
  Returns an individual TabStop object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Position As Single, [Alignment As Variant], [Leader As Variant]) As TabStop`  
  Returns a TabStop object that represents a custom tab stop added to a document.
    - `Position As Single` (required): The position of the tab stop (in points) relative to the left margin.
    - `Alignment As Variant` (optional): The alignment of the tab stop. Can be one of the WdTabAlignment constants.
    - `Leader As Variant` (optional): The type of leader for the tab stop. Can be one of the WdTabLeader constants. If this argument is omitted, wdTabLeaderSpaces is used.
- `ClearAll()`  
  Clears all the custom tab stops from the specified paragraphs.
- `Before(Position As Single) As TabStop`  
  Returns the next TabStop object to the left of Position.
    - `Position As Single` (required): A location on the ruler, in points.
- `After(Position As Single) As TabStop`  
  Returns the next TabStop object to the right of Position.
    - `Position As Single` (required): A location on the ruler, in points.
