# IXmlMap

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002447B-0001-0000-C000-000000000046}  

## Properties (17)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `_Default As HRESULT  (read-only)`
- `Name As HRESULT  (read/write)`
- `IsExportable As HRESULT  (read-only)`
- `ShowImportExportValidationErrors As HRESULT  (read/write)`
- `SaveDataSourceDefinition As HRESULT  (read/write)`
- `AdjustColumnWidth As HRESULT  (read/write)`
- `PreserveColumnFilter As HRESULT  (read/write)`
- `PreserveNumberFormatting As HRESULT  (read/write)`
- `AppendOnImport As HRESULT  (read/write)`
- `RootElementName As HRESULT  (read-only)`
- `RootElementNamespace As HRESULT  (read-only)`
- `Schemas As HRESULT  (read-only)`
- `DataBinding As HRESULT  (read-only)`
- `WorkbookConnection As HRESULT  (read-only)`

## Methods (5)

- `Delete()`
- `Import(Url As String, [Overwrite As Variant], RHS As XlXmlImportResult)`
- `ImportXml(XmlData As String, [Overwrite As Variant], RHS As XlXmlImportResult)`
- `Export(Url As String, [Overwrite As Variant], RHS As XlXmlExportResult)`
- `ExportXml(Data As String, RHS As XlXmlExportResult)`
