# CategoryCollection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244DE-0000-0000-C000-000000000046}  

Represents the collection of visible chart categories in the chart.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CategoryCollection object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the CategoryCollection (returns the number of categories visible in the chart). Read-only.

## Methods (2)

- `Item(Index As Variant) As ChartCategory`  
  Returns a single object from the CategoryCollection object.
    - `Index As Variant` (required): The index number for the object.
- `_Default(Index As Variant) As ChartCategory`
