# Research

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934F7-5A91-11CF-8700-00AA0060263B}  

Provides access to the research service feature of Microsoft PowerPoint.

**Remarks:** The research service feature provides the ability to search multiple custom and third-party references from within PowerPoint. For more information, see the Office 2003 Research Services Software Development Kit (SDK).

## Properties (2)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft PowerPoint application. Read-only.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Research object. Read-only.

## Methods (3)

- `Query(ServiceID As String, [QueryString As Variant], [QueryLanguage As Variant], [UseSelection As Boolean], [LaunchQuery As Boolean])`  
  Specifies a research query.
    - `ServiceID As String` (required): Specifies a GUID that identifies the research service.
    - `QueryString As Variant` (optional): Specifies the query string.
    - `QueryLanguage As Variant` (optional): Specifies the query language of the query string.
    - `UseSelection As Boolean` (optional): True to use the current selection as the query string. This overrides the QueryString parameter if set. Default value is False.
    - `LaunchQuery As Boolean` (optional): True launches the query. False displays the Research task pane scoped to search the specified research service.
- `SetLanguagePair(Language1 As Variant, Language2 As Variant)`  
  Sets the languages for the translation service.
    - `Language1 As Variant` (required): Specifies the language to translate from.
    - `Language2 As Variant` (required): Specifies the language to translate to.
- `IsResearchService(ServiceID As String) As Boolean`  
  Indicates whether the GUID specified in the _ServiceID_ parameter corresponds to a currently configured research service.
    - `ServiceID As String` (required): A GUID that identifies the research service.
