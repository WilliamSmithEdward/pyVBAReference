# SharedResource

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {4DCF0AC2-53CC-45E5-B417-01F3DDD387E5}  

Represents a Microsoft Office theme or image that is available as a shared resource in the database.

**Remarks:** A shared resource is stored once in the database, but can be used many times. For example, you may want to display your company logo on every form that you create. In earlier versions of Access, you had to import the logo into every form. In Access, you can add the logo as a shared image. It will then be displayed in the Image Gallery that appears when you choose the Insert Image menu for the Controls group on the Design tab. To import an image as a SharedResource object, use the AddSharedImage method of the CodeProject object or the AddSharedImage method of the CurrentProject object.

## Properties (3)

- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `Type As AcResourceType  (read-only)`  
  Gets the type of the SharedResource object. Read-only AcResourceType.

## Methods (1)

- `Delete()`  
  Deletes the specified object.
