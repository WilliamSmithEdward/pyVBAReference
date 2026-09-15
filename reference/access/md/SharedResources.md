# SharedResources

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {78B78575-C7B7-4179-824A-95ADBF5423E7}  

Represents the collection of shared resources in the database.

**Remarks:** The SharedResources collection contains Microsoft Office themes and images that are stored once, but used throughout the database. For example, you may want to display your company logo on every form that you create. In earlier versions of Access, you had to import the logo into every form. In Access, you can add the logo as a shared image. It will then be displayed in the Image Gallery that appears when you choose the Insert Image menu for the Controls group on the Design tab. Use the Resources property of the CodeProject object or the Resources property of the CurrentProject object to enumerate the SharedResources collection. To import an image as a SharedResource object, use the AddSharedImage method of the CodeProject object or the AddSharedImage method of the CurrentProject object.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As SharedResource  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only Object.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
