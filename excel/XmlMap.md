# XmlMap

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002447B-0000-0000-C000-000000000046}  

## Properties (17)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `_Default As String  (read-only)`
- `Name As String  (read/write)`
- `IsExportable As Boolean  (read-only)`
- `ShowImportExportValidationErrors As Boolean  (read/write)`
- `SaveDataSourceDefinition As Boolean  (read/write)`
- `AdjustColumnWidth As Boolean  (read/write)`
- `PreserveColumnFilter As Boolean  (read/write)`
- `PreserveNumberFormatting As Boolean  (read/write)`
- `AppendOnImport As Boolean  (read/write)`
- `RootElementName As String  (read-only)`
- `RootElementNamespace As XmlNamespace  (read-only)`
- `Schemas As XmlSchemas  (read-only)`
- `DataBinding As XmlDataBinding  (read-only)`
- `WorkbookConnection As WorkbookConnection  (read-only)`

## Methods (5)

- `Delete()`
- `Import(Url As String, [Overwrite As Variant]) As XlXmlImportResult`
- `ImportXml(XmlData As String, [Overwrite As Variant]) As XlXmlImportResult`
- `Export(Url As String, [Overwrite As Variant]) As XlXmlExportResult`
- `ExportXml(Data As String) As XlXmlExportResult`
