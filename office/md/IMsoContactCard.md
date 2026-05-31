# IMsoContactCard

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03F0-0000-0000-C000-000000000046}  

Represents a contact card.

**Remarks:** You cannot create a new instance of a ContactCard object programmatically. The ContactCard object is returned as an IRibbonControl.Context object when a Microsoft Office Fluent Ribbon callback procedure is triggered from the Contact Card context menu. You use the ContactCard object to determine additional information about the entity displayed in the contact card.

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the IMsoContactCard object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the IMsoContactCard object was created. Read-only.
- `Address As String  (read-only)`  
  Represents the address in a Contact card. Read-only.
- `AddressType As MsoContactCardAddressType  (read-only)`  
  An MsoContactCardAddressType value that represents the address type for the ContactCard object. Read-only.
- `CardType As MsoContactCardType  (read-only)`  
  An MsoContactCardType value that represents the type of contact card. Read-only.
- `Parent As Object  (read-only)`  
  Returns the calling object. Read-only.
