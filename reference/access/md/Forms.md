# Forms

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {B1BB0E80-6128-101B-AF4E-00AA003F0F07}  

The Forms collection contains all the currently open forms in a Microsoft Access database.

**Remarks:** Use the Forms collection in Visual Basic or in an expression to refer to forms that are currently open. For example, you can enumerate the Forms collection to set or return the values of properties of individual forms in the collection. Refer to an individual Form object in the Forms collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific form in the Forms collection, it's better to refer to the form by name because a form's collection index may change. The Forms collection is indexed beginning with zero. If you refer to a form by its index, the first form opened is Forms(0), the second form opened is Forms(1), and so on. If you opened Form1 and then opened Form2, Form2 would be referenced in the Forms collection by its index as Forms(1). If you then closed Form1, Form2 would be referenced in the Forms collection by its index as Forms(0). You can't add or delete a Form object from the Forms collection.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As Form  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only Form.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
