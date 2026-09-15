# FullSeriesCollection

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {288B25A9-98EF-41E5-BEBA-F547D7169BF2}  

Represents the full collection of all chart series, both filtered and unfiltered.

## Properties (4)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of full series in the collection. Read-only.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in PowerPoint, this property returns the hexadecimal number 50575054. Read-only.

## Methods (3)

- `Item(Index As Variant) As Series`  
  Returns an individual Series object from the collection.
    - `Index As Variant` (required): The name or index number of the item to return.
- `_NewEnum() As IUnknown`
- `_Default(Index As Variant) As Series`
