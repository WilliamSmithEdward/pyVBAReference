# SmartTag

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {EF1A0B1D-AD6D-48E6-9905-BEE2A5D38DF9}  

Represents a smart tag that has been added to a control on a form or report. The SmartTag object is a member of the SmartTags collection.

**Remarks:** To return a single SmartTag object, use the Item property of the SmartTags collection, or use SmartTags (index), where index represents the number of the smart tag. To return the collection of actions available for the smart tag, use the SmartTagActions property. To perform a smart tag action, use the Execute method of the SmartTagAction object.

## Properties (7)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Name As String  (read-only)`  
  Use the Name property to determine the string expression that identifies the name of an object. Read-only String.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Properties As _SmartTagProperties  (read-only)`  
  Returns a SmartTagProperties collection that represents the collection of properties for a particular smart tag. Read-only.
- `SmartTagActions As _SmartTagActions  (read-only)`  
  Returns a SmartTagActions collection that represents the actions available for a specific smart tag. Read-only.
- `XML As String  (read-only)`  
  Returns a String that represents the related XML code for a smart tag. Read-only.
- `IsMissing As Boolean  (read-only)`  
  Returns True if the specified smart tag is not installed or is installed incorrectly. Read-only Boolean.

## Methods (1)

- `Delete()`  
  Deletes the specified object.
