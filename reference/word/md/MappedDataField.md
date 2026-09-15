# MappedDataField

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {5D311669-EA51-11D3-87CC-00105AA31A34}  

A mapped data field is a field contained within Microsoft Word that represents commonly used name or address information, such as "First Name." If a data source contains a "First Name" field or a variation (such as "First_Name," "FirstName," "First," or "FName"), the field in the data source will automatically map to the corresponding mapped data field in Word. If a document or template is to be merged with more than one data source, mapped data fields make it unnecessary to reenter the fields into the document to agree with the field names in the database.

**Remarks:** Use the MappedDataFields property to return a MappedDataField object. This example returns the data source field name for the wdFirstName mapped data field. This example assumes the current document is a mail merge document. A blank string value returned for the DataFieldName property indicates that the mapped data field is not mapped to a field in the data source.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MappedDataField object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `DataFieldName As String  (read-only)`  
  Sets or returns a String that represents the name of the field in the mail merge data source to which a mapped data field maps. Read/write.
- `Name As String  (read-only)`  
  Returns name of the specified object. Read-only String.
- `Value As String  (read-only)`  
  Returns the contents of the mail merge data field or mapped data field for the current record. Read-only String.
- `DataFieldIndex As Long  (read/write)`  
  Returns or sets a Long that represents the corresponding field number in the mail merge data source to which a mapped data field maps. Read/write.
