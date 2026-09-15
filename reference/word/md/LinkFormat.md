# LinkFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020931-0000-0000-C000-000000000046}  

Represents the linking characteristics for an OLE object or picture.

**Remarks:** Use the LinkFormat property for a shape, inline shape, or field to return the LinkFormat object. The following example breaks the link for the first shape on the active document. Not all types of shapes, inline shapes, and fields can be linked to a source. Use the Type property for the Shape and InlineShape objects to determine whether a particular shape can be linked. The Type property for a Field object returns the type of field. Use both the Update method and the AutoUpdate property to update links. To return or set the full path for a particular link's source file, use the SourceFullName property.

## Properties (10)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified LinkFormat object.
- `AutoUpdate As Boolean  (read/write)`  
  True if the specified link is updated automatically when the container file is opened or when the source file is changed. Read/write Boolean.
- `SourceName As String  (read-only)`  
  Returns the name of the source file for the specified linked OLE object, picture, or field. Read-only String.
- `SourcePath As String  (read-only)`  
  Returns the path of the source file for the specified linked OLE object, picture, or field. Read-only String.
- `Locked As Boolean  (read/write)`  
  True if a Field, InlineShape, or Shape object is locked to prevent automatic updating. Read/write Boolean.
- `Type As WdLinkType  (read-only)`  
  Returns the link type. Read-only WdLinkType.
- `SourceFullName As String  (read/write)`  
  Returns or sets the path and name of the source file for the specified linked OLE object, picture, or field. Read/write String.
- `SavePictureWithDocument As Boolean  (read/write)`  
  True if the specified picture is saved with the document. Read/write Boolean.

## Methods (2)

- `BreakLink()`  
  Breaks the link between the source file and the specified OLE object, picture, or linked field.
- `Update()`  
  Updates the specified link format.
