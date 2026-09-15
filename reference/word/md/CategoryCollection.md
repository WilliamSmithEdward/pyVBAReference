# CategoryCollection

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {04124C2D-039D-4442-9C68-8FA38D11DDD6}  

Represents the collection of visible chart categories in the document.

**Remarks:** Categories are visible if they have not been filtered out of a chart.

## Properties (4)

- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CategoryCollection object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of categories in the collection. Read-only.
- `Application As Object  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (3)

- `Item(Index As Variant) As ChartCategory`  
  Returns an individual chart category.
    - `Index As Variant` (required): The name or index number for the object.
- `_NewEnum() As IUnknown`
- `_Default(Index As Variant) As ChartCategory`
