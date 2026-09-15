# RepeatingSectionItemColl

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {53FACA33-DB22-473F-BB51-96C2C86C9304}  

Represents the collection of RepeatingSectionItem objects.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified RepeatingSectionItemColl object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of repeating section items in the collection. Read-only.

## Methods (1)

- `Item(Index As Long) As RepeatingSectionItem`  
  Returns an individual repeating section item.
    - `Index As Long` (required): The index number of the item in the collection.
