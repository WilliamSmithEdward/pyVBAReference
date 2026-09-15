# SmartTagAction

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {0D944D89-82BC-43DE-9659-699DD3FBCD72}  

Represents a single action for a smart tag.

**Remarks:** Smart tag actions are processes that are programmed into smart tags that allow users to perform certain functions related to the smart tag. For example, one action for a smart tag might be to access a website, while another action inserts contact information from Microsoft Outlook, while yet another action displays a map and driving directions. To perform an action represented by a SmartTagAction object, use the Execute method. To return the parent control of the SmartTag object, use the Parent property.

## Properties (3)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Use the Name property to determine the string expression that identifies the name of an object. Read-only String.

## Methods (1)

- `Execute()`  
  The Execute method performs the specified smart tag action.
