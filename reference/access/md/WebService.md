# WebService

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {7A7B2B33-A93D-4E04-BFC7-6041AAA0046D}  

Represents a Data Service data connection.

**Remarks:** Use the Item property of the WebServices collection to return a WebService object. A Data Service data connection may contain one or more entities. Each entity specifies an external content type. Used throughout the functionality and services offered by Business Connectivity Services, external content types are reusable metadata descriptions of connectivity information and data definitions plus the behaviors that you want to apply to a certain category of external data. Use the Entities property to return the entities defined for a Data Service data connection.

## Properties (3)

- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `Entities As Entities  (read-only)`  
  Gets the collection of entities defined for the specified Data Service data connection. Read-only Entities.
