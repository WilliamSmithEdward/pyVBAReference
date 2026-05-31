# RulerLevels2

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C2-0000-0000-C000-000000000046}  

A collection of all the RulerLevel2 objects on the specified ruler.

**Remarks:** Each RulerLevel2 object represents the first-line and left indent for text at a particular outline level. This collection always contains five members-one for each of the available outline levels.

**Example:**

```vba
With ActivePresentation.SlideMaster.TextStyles(ppBodyStyle).Ruler2
 .Levels(1).FirstMargin = 0
 .Levels(1).LeftMargin = 40
 .Levels(2).FirstMargin = 60
 .Levels(2).LeftMargin = 100
 .Levels(3).FirstMargin = 120
 .Levels(3).LeftMargin = 160
 .Levels(4).FirstMargin = 180
 .Levels(4).LeftMargin = 220
 .Levels(5).FirstMargin = 240
 .Levels(5).LeftMargin = 280
End With
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the RulerLevels2 object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the RulerLevels2 object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the RulerLevels2 object. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the RulerLevels2 collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As Variant) As RulerLevel2`  
  Gets a member of the RulerLevels2 collection.
    - `Index As Variant` (required): The index number of the object to be returned.
