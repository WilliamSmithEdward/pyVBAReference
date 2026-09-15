# AccessObjectProperty

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {1FE3E471-A7D0-11D1-9944-006008197D41}  

An AccessObjectProperty object represents a built-in or user-defined characteristic of an AccessObject object.

**Remarks:** Every AccessObject object contains an AccessObjectProperties collection that has AccessObjectProperty objects corresponding to the properties of that AccessObject object. The user can also define AccessObjectProperty objects and append them to the AccessObjectProperties collection of some AccessObject objects. You can create user-defined properties for the following objects: - CodeData, CodeProject, CurrentProject, and CurrentData objects - AccessObject objects in the following collections: - CurrentProject and CodeProject object collections: - AllForms - AllReports - AllMacros - AllModules - AllTables - CodeData and CodeProject object collections: - AllQueries - AllViews - AllStoredProcedures - AllDatabaseDiagrams To add a user-defined property, use the Add method to create and add an AccessObjectProperty object with a unique Name property and Value property. The object to which you are adding the user-defined property must already be appended to a collection.

## Properties (2)

- `Name As String  (read-only)`  
  Use the Name property to determine the string expression that identifies the name of an object. Read-only String.
- `Value As Variant  (read/write)`  
  Determines or specifies the value of a built-in property of an AccessObject object. Read/write Variant.
