# FileConverters

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002099A-0000-0000-C000-000000000046}  

A collection of FileConverter objects that represent all the file converters available for opening and saving files.

**Remarks:** Use the FileConverters property to return the FileConverters collection. The following example determines whether a WordPerfect 6.0 converter is available. The Add method isn't available for the FileConverters collection. FileConverter objects are added during installation of Microsoft Office or by installing supplemental converters. Use FileConverters (Index), where Index is a class name or index number, to return a single FileConverter object. The following example displays the extensions associated with the Microsoft Excel worksheet converter. The index number represents the position of the file converter in the FileConverters collection. The following example displays the format name of the first file converter. File converters for saving documents are listed in the Save As dialog box. File converters for opening documents appear in a dialog box if the Confirm conversion at Open check box is selected on the General tab in the Options dialog box (Tools menu).

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FileConverters object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of file converters in the collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `ConvertMacWordChevrons As WdChevronConvertRule  (read/write)`  
  Controls whether text enclosed in chevron characters (" ") is converted to merge fields. Read/write Long. .

## Methods (1)

- `Item(Index As Variant) As FileConverter`  
  Returns an individual FileConverter object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
