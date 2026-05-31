# IModel

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244DB-0001-0000-C000-000000000046}  

## Properties (16)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `ModelTables As HRESULT  (read-only)`
- `ModelRelationships As HRESULT  (read-only)`
- `DataModelConnection As HRESULT  (read-only)`
- `Name As HRESULT  (read-only)`
- `ModelMeasures As HRESULT  (read-only)`
- `ModelFormatGeneral As HRESULT  (read-only)`
- `ModelFormatDate As HRESULT  (read-only)`
- `ModelFormatDecimalNumber As HRESULT  (read-only)`
- `ModelFormatWholeNumber As HRESULT  (read-only)`
- `ModelFormatPercentageNumber As HRESULT  (read-only)`
- `ModelFormatScientificNumber As HRESULT  (read-only)`
- `ModelFormatCurrency As HRESULT  (read-only)`
- `ModelFormatBoolean As HRESULT  (read-only)`

## Methods (4)

- `Refresh()`
- `AddConnection(ConnectionToDataSource As WorkbookConnection, RHS As WorkbookConnection)`
- `CreateModelWorkbookConnection(ModelTable As Variant, RHS As WorkbookConnection)`
- `Initialize()`
