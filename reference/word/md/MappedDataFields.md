# MappedDataFields

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {799A6814-EA41-11D3-87CC-00105AA31A34}  

A collection of MappedDataField objects that represents all the mapped data fields available in Microsoft Word.

**Remarks:** Use the MappedDataFields property of the MailMergeDataSource object to return the MappedDataFields collection. This example creates a tabbed list of the mapped data fields available in Word and the fields in the data source to which they are mapped. This example assumes that the current document is a mail merge document and that the data source fields have corresponding mapped data fields.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MappedDataFields object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of mapped data fields in the collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As WdMappedDataFields) As MappedDataField`  
  Returns an individual MappedDataField object.
    - `Index As WdMappedDataFields` (required): The specified mapped data field.
