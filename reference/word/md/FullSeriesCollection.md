# FullSeriesCollection

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {4DACC469-630B-457E-9C8F-08158D57FC7C}  

Represents the full collection of all chart series, both filtered and unfiltered.

## Properties (4)

- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FullSeriesCollection object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of full series in the collection. Read-only.
- `Application As Object  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (3)

- `Item(Index As Variant) As Series`  
  Returns an individual full series item from the collection.
    - `Index As Variant` (required): The name or index number of the item to return.
- `_NewEnum() As IUnknown`
- `_Default(Index As Variant) As Series`
