# Reviewer

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {47CEF4AE-DC32-4220-8AA5-19CCC0E6633A}  

Represents a single reviewer of a document in which changes have been tracked. The Reviewer object is a member of the Reviewers collection.

**Remarks:** Use Reviewers (_index_), where _index_ is the name or number of the reviewer, to return a Reviewer object. Use the Visible property to display or hide individual reviewers in a document. The following code example hides the reviewer named "Jeff Smith" and displays the reviewer named "Judy Lew." This assumes that "Jeff Smith" and "Judy Lew" are members of the Reviewers collection. If they are not, you will receive an error.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Reviewer object.
- `Visible As Boolean  (read/write)`  
  True if the specified object is visible. Read/write Boolean.
