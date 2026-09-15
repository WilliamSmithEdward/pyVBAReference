# Entity

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3D1E34BD-F23F-46D1-B80A-2BA8ACA850A9}  

Represents an entity defined in a Data Service data connection.

**Remarks:** Use the Item property of the Entities collection to return an Entity object. Use the Operations property to return the operations defined for the specified entity. A Data Service data connection may contain one or more entities. Each entity specifies an external content type. Used throughout the functionality and services offered by Business Connectivity Services, external content types are reusable metadata descriptions of connectivity information and data definitions plus the behaviors that you want to apply to a certain category of external data. For more information about external content types, see What Are External Content Types.

## Properties (3)

- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `Operations As Operations  (read-only)`  
  Gets the collection of operations defined for the specified Entity object. Read-only Operations.
