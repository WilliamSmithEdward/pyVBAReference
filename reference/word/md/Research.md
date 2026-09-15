# Research

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {E6AAEC05-E543-4085-BA92-9BF7D2474F51}  

Provides access to the research service feature of Microsoft Word.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Research object.
- `FavoriteService As String  (read/write)`  
  Returns or sets a String that specifies the favorite research service.

## Methods (3)

- `Query(ServiceID As String, [QueryString As String], [QueryLanguage As WdLanguageID], [UseSelection As Boolean], [LaunchQuery As Boolean]) As Variant`  
  Specifies a research query.
    - `ServiceID As String` (required): Specifies a GUID that identifies the research service.
    - `QueryString As String` (optional): Specifies the query string.
    - `QueryLanguage As WdLanguageID` (optional): Specifies the query language of the query string.
    - `UseSelection As Boolean` (optional): True to use the current selection as the query string. This overrides the QueryString parameter if set. Default value is False.
    - `LaunchQuery As Boolean` (optional): True launches the query. False displays the Research task pane scoped to search the specified research service.
- `SetLanguagePair(LanguageFrom As WdLanguageID, LanguageTo As WdLanguageID) As Variant`  
  Sets the languages for the translation service.
    - `LanguageFrom As WdLanguageID` (required): Specifies the language to translate from.
    - `LanguageTo As WdLanguageID` (required): Specifies the language to translate to.
- `IsResearchService(ServiceID As String) As Boolean`  
  Indicates whether the GUID specified in the ServiceID parameter corresponds to a currently configured service.
    - `ServiceID As String` (required): Specifies a GUID that identifies the research service.
