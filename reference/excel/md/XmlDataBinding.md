# XmlDataBinding

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024478-0000-0000-C000-000000000046}  

Represents the connection to the source data for an XmlMap object.

**Remarks:** Use the ClearSettings method to remove a data binding. Use the LoadSettings method to initialize the settings for an XmlDataBinding object. Use the Refresh method to refresh a data binding.

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As String  (read-only)`
- `SourceUrl As String  (read-only)`  
  Returns a String that represents the path to the XML data file or the web service that provides the source data for the specified data binding. Read-only.

## Methods (3)

- `Refresh() As XlXmlImportResult`  
  Retrieves XML data by using the current connection settings of the specified XmlDataBinding object.
- `LoadSettings(Url As String)`  
  Initializes the specified data binding with settings from an XML data file or a Data Retrieval Service Connection (.uxdc) file.
    - `Url As String` (required): The path to the XML data file. The path is specified in the Uniform Resource Locator (URL) or universal naming convention (UNC) format.
- `ClearSettings()`  
  Removes the specified data binding.
