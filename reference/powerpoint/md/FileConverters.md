# FileConverters

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A50-F07E-4CA4-AF6F-BEF486AA4E6F}  

A collection of FileConverter objects that represent all the file converters available for opening and saving files.

**Remarks:** Use the FileConverters property to return the FileConverters collection. The following example determines whether a WordPerfect 6.0 converter is available. The Add method isn't available for the FileConverters collection. FileConverter objects are added during installation of Microsoft Office or by installing supplemental converters. Use FileConverters (Index), where Index is a class name or index number, to return a single FileConverter object. The following example displays the extensions associated with the Microsoft Excel worksheet converter. The index number represents the position of the file converter in the FileConverters collection. The following example displays the format name of the first file converter. File converters for saving documents are listed in the Save As dialog box. File converters for opening documents appear in a dialog box if the Confirm conversion at Open check box is selected on the General tab in the Options dialog box.

## Properties (1)

- `Count As Long  (read-only)`  
  Returns a Long that represents the number of file converters in the collection. Read-only.

## Methods (1)

- `Item(Index As Variant) As FileConverter`  
  Returns an individual FileConverter object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
