# MailMergeDataSource

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002091D-0000-0000-C000-000000000046}  

Represents the mail merge data source in a mail merge operation.

**Remarks:** Use the DataSource property to return the MailMergeDataSource object. The following example displays the name of the data source associated with the active document. The following example displays the field names in the data source associated with the active document. The following example opens the data source associated with Form letter.doc and determines whether the FirstName field includes the name "Kate."

## Properties (20)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MailMergeDataSource object.
- `Name As String  (read-only)`  
  Returns name of the specified object. Read-only String.
- `HeaderSourceName As String  (read-only)`  
  Returns the path and file name of the header source attached to the specified mail merge main document. Read-only String.
- `Type As WdMailMergeDataSource  (read-only)`  
  Returns the type of mail merge data source. Read-only WdMailMergeDataSource.
- `HeaderSourceType As WdMailMergeDataSource  (read-only)`  
  Returns a value that indicates the way the header source is being supplied for the mail merge operation. Read-only WdMailMergeDataSource.
- `ConnectString As String  (read-only)`  
  Returns the connection string for the specified mail merge data source. Read-only String.
- `QueryString As String  (read/write)`  
  Returns or sets the query string (SQL statement) used to retrieve a subset of the data in a mail merge data source. Read/write String.
- `ActiveRecord As WdMailMergeActiveRecord  (read/write)`  
  Returns or sets the active mail merge record. Can be either a valid record number in the query result or one of the WdMailMergeActiveRecord constants.
- `FirstRecord As Long  (read/write)`  
  Returns or sets the number of the first record to be merged in a mail merge operation. Read/write Long.
- `LastRecord As Long  (read/write)`  
  Returns or sets the number of the last record to be merged in a mail merge operation. Read/write Long.
- `FieldNames As MailMergeFieldNames  (read-only)`  
  Returns a MailMergeFieldNames collection that represents the names of all the fields in the specified mail merge data source. Read-only.
- `DataFields As MailMergeDataFields  (read-only)`  
  Returns a MailMergeDataFields collection that represents the fields in the specified mail merge data source. Read-only.
- `RecordCount As Long  (read-only)`  
  Returns a Long that represents the number of records in the data source. Read-only.
- `Included As Boolean  (read/write)`  
  True if a record is included in a mail merge. Read/write Boolean.
- `InvalidAddress As Boolean  (read/write)`  
  True for Microsoft Word to mark a record in a mail merge data source if it contains invalid data in an address field. Read/write Boolean.
- `InvalidComments As String  (read/write)`  
  If the InvalidAddress property is True, returns or sets a String that describes an invalid address error. Read/write.
- `MappedDataFields As MappedDataFields  (read-only)`  
  Returns a MappedDataFields collection that represents the mapped data fields available in Microsoft Word.
- `TableName As String  (read-only)`  
  Returns a String with the SQL query used to retrieve the records from the data source file attached to a mail merge document. Read-only.

## Methods (4)

- `FindRecord(FindText As String, [Field As Variant]) As Boolean`  
  Searches the contents of the specified mail merge data source for text in a particular field. Returns True if the search text is found. Boolean.
    - `FindText As String` (required): The text to be looked for.
    - `Field As Variant` (optional): The name of the field to be searched.
- `SetAllIncludedFlags(Included As Boolean)`  
  Includes or excludes flagged records in a data source from a mail merge.
    - `Included As Boolean` (required): True to include all data source records in a mail merge. False to exclude all data source records from a mail merge.
- `SetAllErrorFlags(Invalid As Boolean, InvalidComment As String)`  
  Marks all records in a mail merge data source as containing invalid data in an address field.
    - `Invalid As Boolean` (required): True marks all records in the data source of a mail merge as invalid.
    - `InvalidComment As String` (required): Text describing the invalid setting.
- `Close()`  
  Closes the specified Mail Merge data source.
