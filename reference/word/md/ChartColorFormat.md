# ChartColorFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {DD8F80B8-9B80-4E89-9BEC-F12DF35E43B3}  

Represents the color of a one-color object or the foreground or background color of an object with a gradient or patterned fill.

## Properties (7)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `SchemeColor As Long  (read/write)`  
  Returns or sets the index of a color in the current color scheme. Read/write Long.
- `RGB As Long  (read-only)`  
  Returns the red-green-blue value of the specified color. Read-only Long.
- `_Default As Long  (read-only)`
- `Type As Long  (read-only)`  
  Returns the color format type. Read-only Long.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
