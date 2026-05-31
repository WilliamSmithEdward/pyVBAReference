# Model

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244DB-0000-0000-C000-000000000046}  

## Properties (16)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified object. Read-only.
- `ModelTables As ModelTables  (read-only)`  
  Returns a ModelTables object that represents a collection of tables inside the data model. Read-only.
- `ModelRelationships As ModelRelationships  (read-only)`  
  Returns a ModelRelationships object that represents the collection of relationships between data model tables. Read-only.
- `DataModelConnection As WorkbookConnection  (read-only)`  
  Returns the model WorkbookConnection object from the workbook connections collection that connects to the model.
- `Name As String  (read-only)`  
  Returns a String value representing the name of the Model object. Read-only.
- `ModelMeasures As ModelMeasures  (read-only)`  
  Returns a ModelMeasures object that represents the collection of model measures in the data model. Read-only.
- `ModelFormatGeneral As ModelFormatGeneral  (read-only)`  
  Returns a ModelFormatGeneral object that represents formatting of type general in the data model. Read-only.
- `ModelFormatDate As ModelFormatDate  (read-only)`  
  Returns a ModelFormatDate object that represents formatting of type date in the data model. Read-only.
- `ModelFormatDecimalNumber As ModelFormatDecimalNumber  (read-only)`  
  Returns a ModelFormatDecimalNumber object that represents formatting of type decimal number in the data model. Read-only.
- `ModelFormatWholeNumber As ModelFormatWholeNumber  (read-only)`  
  Returns a ModelFormatWholeNumber object that represents formatting of type whole number in the data model. Read-only.
- `ModelFormatPercentageNumber As ModelFormatPercentageNumber  (read-only)`  
  Returns a ModelFormatPercentageNumber object that represents formatting of type percentage number in the data model. Read-only.
- `ModelFormatScientificNumber As ModelFormatScientificNumber  (read-only)`  
  Returns a ModelFormatScientificNumber object that represents formatting of type scientific number in the data model. Read-only.
- `ModelFormatCurrency As ModelFormatCurrency  (read-only)`  
  Returns a ModelFormatCurrency object that represents formatting of type currency in the data model. Read-only.
- `ModelFormatBoolean As ModelFormatBoolean  (read-only)`  
  Returns a ModelFormatBoolean object that represents formatting of type True/False in the data model. Read-only.

## Methods (4)

- `Refresh()`  
  Refreshes all data sources associated with the model, fully reprocesses the model, and updates all Excel data features associated with the model.
- `AddConnection(ConnectionToDataSource As WorkbookConnection) As WorkbookConnection`  
  Adds a new WorkbookConnection to the model with the same properties as the one supplied as an argument.
    - `ConnectionToDataSource As WorkbookConnection` (required): The Workbook connection.
- `CreateModelWorkbookConnection(ModelTable As Variant) As WorkbookConnection`  
  Returns a WorkbookConnection object of type ModelConnection.
    - `ModelTable As Variant` (required): Either a model table name or a model table object.
- `Initialize()`  
  Initializes the Workbook's data model. This is called by default the first time the model is used.
