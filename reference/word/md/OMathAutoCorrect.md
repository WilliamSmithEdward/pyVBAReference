# OMathAutoCorrect

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {6F9D1F68-06F7-49EF-8902-185E54EB5E87}  

Represents the math AutoCorrect feature in Microsoft Word. To access the math AutoCorrect entries, use the OMathAutoCorrectEntries collection.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathAutoCorrect object.
- `ReplaceText As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Microsoft Word automatically replaces strings in equations with the corresponding math AutoCorrect definitions. Read/write.
- `UseOutsideOMath As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Microsoft Word uses math autocorrect rules outside equations in a document. Read/write.
- `Entries As OMathAutoCorrectEntries  (read-only)`  
  Returns an OMathAutoCorrectEntries collection that represents the collection of equation autocorrect entries within the equation autocorrect feature. Read-only.
- `Functions As OMathRecognizedFunctions  (read-only)`  
  Returns an OMathRecognizedFunctions collection that represents the recognized functions that are automatically corrected using the equation autocorrect feature. Read-only.
