# Entities

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {4B0E334D-B734-458A-A041-B528D031D4E5}  

Represents the collection of entities defined in a Data Service data connection.

**Remarks:** A Data Service data connection may contain one or more entities. Each entity specifies an external content type. Used throughout the functionality and services offered by Business Connectivity Services, external content types are reusable metadata descriptions of connectivity information and data definitions plus the behaviors that you want to apply to a certain category of external data. Use the Entities property to return the entities defined for a Data Service data connection. Use the Item property to return an Entity object. For more information about external content types, see What Are External Content Types.

## Properties (3)

- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As Entity  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only Object.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
