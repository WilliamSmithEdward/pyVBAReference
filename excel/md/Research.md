# Research

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244AC-0000-0000-C000-000000000046}  

Represents the controls of a Research query.

**Remarks:** When working with Research queries, you must have an existing GUID that corresponds to a live data source. If the data source is unavailable or does not exist, a run-time error occurs.

**Example:**

```vba
Worksheets("Sheet1").Research.Translate = True
```

## Properties (3)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.

## Methods (3)

- `Query(ServiceID As String, [QueryString As Variant], [QueryLanguage As Variant], [UseSelection As Variant], [LaunchQuery As Variant]) As Variant`  
  Specifies a research query.
    - `ServiceID As String` (required): Specifies a GUID that identifies the research service.
    - `QueryString As Variant` (optional): Specifies the query string.
    - `QueryLanguage As Variant` (optional): Specifies the query language of the query string.
    - `UseSelection As Variant` (optional): True to use the current selection as the query string. This overrides the _QueryString_ parameter if set. Default value is False.
    - `LaunchQuery As Variant` (optional): True launches the query. False displays the Research task pane scoped to search the specified research service.
- `IsResearchService(ServiceID As String) As Boolean`  
  Indicates whether the GUID specified in the _ServiceID_ parameter corresponds to a currently configured service.
    - `ServiceID As String` (required): Specifies a GUID that identifies the research service.
- `SetLanguagePair(LanguageFrom As Long, LanguageTo As Long) As Variant`  
  Sets the languages for the translation service.
    - `LanguageFrom As Long` (required): Specifies the language to translate from.
    - `LanguageTo As Long` (required): Specifies the language to translate to.
