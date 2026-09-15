# SmartTagActions

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {AA533187-6399-4E6C-B6EC-6FC999E1C855}  

Represents a collection of actions for an individual smart tag or a type of smart tag.

**Remarks:** Smart tag actions are processes that are programmed into smart tags; they allow users to perform certain functions related to the smart tag. For example, one action for a smart tag might be to access a website, while another action inserts contact information from Microsoft Outlook, while yet another action displays a map and driving directions. To return the collection of actions related to a smart tag, use the SmartTagActions property of the SmartTag object.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As _SmartTagAction  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only SmartTagAction.
