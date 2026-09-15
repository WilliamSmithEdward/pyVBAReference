# Guide

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {3D2F865B-E2DB-4896-BC35-6A006DF896DC}  

Represents a drawing guide in the presentation or custom layout.

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Orientation As PpGuideOrientation  (read-only)`  
  Returns the orientation of the drawing guide (horizontal or vertical) as a constant from the PpGuideOrientation enumeration. Read-only.
- `Position As Single  (read/write)`  
  Returns or sets a Single that represents the position of the drawing guide along the x- or y-axis. Read/write.
- `Color As ColorFormat  (read-only)`  
  Returns a ColorFormat object that represents the color of the drawing guide. Read-only.

## Methods (1)

- `Delete()`  
  Removes the drawing guide from the presentation or custom layout.
