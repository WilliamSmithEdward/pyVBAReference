# Guides

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {1641E775-2277-46DE-A06D-8C49C3C5D5E7}  

A collection of Guide objects in a presentation or custom layout.

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.

## Methods (2)

- `Item(Index As Long) As Guide`  
  Returns a Guide object from the collection.
- `Add(Orientation As PpGuideOrientation, Position As Single) As Guide`  
  Adds a new drawing guide to the presentation, slide master, or custom layout.
