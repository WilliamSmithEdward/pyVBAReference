# TabStop

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020954-0000-0000-C000-000000000046}  

Represents a single tab stop. The TabStop object is a member of the TabStops collection. The TabStops collection represents all the custom and default tab stops in a paragraph or group of paragraphs.

**Remarks:** Use TabStops (Index), where Index is the location of the tab stop (in points) or the index number, to return a single TabStop object. Tab stops are indexed numerically from left to right along the ruler. The following example removes the first custom tab stop from the selected paragraphs. The following example adds a right-aligned tab stop positioned at 2 inches to the selected paragraphs. Use the Add method to add a tab stop. The following example adds two tab stops to the selected paragraphs. The first tab stop is a left-aligned tab with a dotted tab leader positioned at 1 inch (72 points). The second tab stop is centered and is positioned at 2 inches. You can also add a tab stop by specifying a location with the TabStops property. The following example adds a right-aligned tab stop positioned at 2 inches to the selected paragraphs.

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TabStop object.
- `Alignment As WdTabAlignment  (read/write)`  
  Returns or sets a WdTabAlignment constant that represents the alignment for the specified tab stop. Read/write.
- `Leader As WdTabLeader  (read/write)`  
  Returns or sets the leader for the specified TabStop object. Read/write WdTabLeader.
- `Position As Single  (read/write)`  
  Returns or sets the position of a tab stop relative to the left margin. Read/write Single.
- `CustomTab As Boolean  (read-only)`  
  True if the specified tab stop is a custom tab stop. Read-only Boolean.
- `Next As TabStop  (read-only)`  
  Returns the next tabstop in the collection. Read-only.
- `Previous As TabStop  (read-only)`  
  Returns the previous tab stop in the collection. Read-only.

## Methods (1)

- `Clear()`  
  Removes the specified custom tab stop.
