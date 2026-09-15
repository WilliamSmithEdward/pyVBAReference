# CategoryCollection

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {2432F529-514B-4575-AA71-1754C74A13D6}  

Represents the collection of visible chart categories in the presentation.

**Remarks:** Categories are visible if they have not been filtered out of a chart.

## Properties (4)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of categories in the collection. Read-only.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in PowerPoint, this property returns the hexadecimal number 50575054. Read-only.

## Methods (2)

- `Item(Index As Variant) As ChartCategory`  
  Returns an individual chart category.
- `_Default(Index As Variant) As ChartCategory`
