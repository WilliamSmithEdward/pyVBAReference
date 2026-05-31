# XmlMap

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002447B-0000-0000-C000-000000000046}  

Represents an XML map that has been added to a workbook.

**Remarks:** Use the Add method of the XmlMaps collection to add an XML map to a workbook. Use the Import method to import XML data from an XML data file into cells mapped to the specified XmlMap. The ImportXml method imports XML data for a String variable. Use the Export method to export data from cells mapped to the specified XmlMap. The ExportXml method exports XML data to a String variable.

## Properties (17)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As String  (read-only)`
- `Name As String  (read/write)`  
  Returns or sets a String value that represents the friendly name used to uniquely identify a mapping in the workbook.
- `IsExportable As Boolean  (read-only)`  
  Returns True if Microsoft Excel can use the XPath objects in the specified schema map to export XML data, and all XML lists mapped to the specified schema map can be exported.
- `ShowImportExportValidationErrors As Boolean  (read/write)`  
  Returns or sets whether to display a dialog box that details schema-validation errors when data is imported or exported through the specified XML schema map. The default value is False. Read/write Boolean.
- `SaveDataSourceDefinition As Boolean  (read/write)`  
  True if the data source definition of the specified XML schema map is saved with the workbook. The default value is True. Read/write Boolean.
- `AdjustColumnWidth As Boolean  (read/write)`  
  True if the column widths are automatically adjusted for the best fit each time that you refresh the specified XML map.
- `PreserveColumnFilter As Boolean  (read/write)`  
  Returns or sets whether filtering is preserved when the specified XML map is refreshed. Read/write Boolean.
- `PreserveNumberFormatting As Boolean  (read/write)`  
  True if number formatting on cells mapped to the specified XML schema map are preserved when the schema map is refreshed. The default value is False. Read/write Boolean.
- `AppendOnImport As Boolean  (read/write)`  
  True if you want to append new rows to XML lists that are bound to the specified schema map when you are importing new data or refreshing an existing connection.
- `RootElementName As String  (read-only)`  
  Returns a String that represents the name of the root element for the specified XML schema map. Read-only.
- `RootElementNamespace As XmlNamespace  (read-only)`  
  Returns an XmlNamespace object that represents the root element for the specified XML schema map. Read-only.
- `Schemas As XmlSchemas  (read-only)`  
  Returns an XmlSchemas collection that represents the schemas that the specified XmlMap object contains. Read-only.
- `DataBinding As XmlDataBinding  (read-only)`  
  Returns an XmlDataBinding object that represents the binding associated with the specified schema map. Read-only.
- `WorkbookConnection As WorkbookConnection  (read-only)`  
  Returns a new connection for the specified XMLMap object. Read-only.

## Methods (5)

- `Delete()`  
  Removes the specified XML map from the workbook.
- `Import(Url As String, [Overwrite As Variant]) As XlXmlImportResult`  
  Imports data from the specified XML data file into cells that have been mapped to the specified XmlMap object.
    - `Url As String` (required): The path to the XML data to import. The path can be specified in Universal Naming convention (UNC) or Uniform Resource Locator (URL) format. The file can be an XML data file.
    - `Overwrite As Variant` (optional): Set to True to overwrite existing data. Set to False to append to existing data. The default value is False.
- `ImportXml(XmlData As String, [Overwrite As Variant]) As XlXmlImportResult`  
  Imports XML data from a String variable into cells that have been mapped to the specified XmlMap object.
    - `XmlData As String` (required): The string that contains the XML data to import.
    - `Overwrite As Variant` (optional): Specifies whether to overwrite the contents of cells that are currently mapped to the specified XML map. Set to True to overwrite the cells; set to False to append the data to the existing range.
- `Export(Url As String, [Overwrite As Variant]) As XlXmlExportResult`  
  Exports the contents of cells mapped to the specified XmlMap object to an XML data file.
    - `Url As String` (required): The path and file name of the XML data file to export to.
    - `Overwrite As Variant` (optional): Set to True to overwrite the file specified in the _Url_ parameter if the file exists. The default value is False.
- `ExportXml(Data As String) As XlXmlExportResult`  
  Exports the contents of cells mapped to the specified XmlMap object to a String variable.
    - `Data As String` (required): The variable to export the data to.
