# ContactCard

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03F1-0000-0000-C000-000000000046}  

Represents a Microsoft Office contact card.

**Remarks:** You can create a new instance of a ContactCard object by calling the Microsoft Outlook NameSpace.CreateContactCard method. Before you attempt to do so, you must sign in to a Microsoft Outlook session.

## Properties (2)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the parent Office application for the ContactCard object. Read-only.
- `Creator As Long  (read-only)`  
  Returns a Long that indicates the application in which the ContactCard object was created. Read-only.

## Methods (2)

- `Close()`  
  Closes the contact card.
- `Show(CardStyle As MsoContactCardStyle, RectangleLeft As Long, RectangleRight As Long, RectangleTop As Long, RectangleBottom As Long, HorizontalPosition As Long, [ShowWithDelay As Boolean])`  
  Displays the contact card at the specified _x_-coordinate position outside the specified rectangle.
