# SmartTags

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {73778F0A-9743-4DF3-BBFA-941712488FEA}  

Represents the collection of smart tags for a control on a form, report, or data access page.

**Remarks:** To return a single SmartTag object, use the Item property or use SmartTags (Index), where Index represents the number of the smart tag.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As _SmartTag  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only SmartTag.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.

## Methods (1)

- `Add(Name As String) As _SmartTag`  
  Adds a smart tag to a form or control.
    - `Name As String` (required): The name of the smart tag to add.
