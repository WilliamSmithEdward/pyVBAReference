# FileSearch

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0332-0000-0000-C000-000000000046}  

## Properties (15)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `SearchSubFolders As Boolean  (read/write)`
- `MatchTextExactly As Boolean  (read/write)`
- `MatchAllWordForms As Boolean  (read/write)`
- `FileName As String  (read/write)`
- `FileType As MsoFileType  (read/write)`
- `LastModified As MsoLastModified  (read/write)`
- `TextOrProperty As String  (read/write)`
- `LookIn As String  (read/write)`
- `FoundFiles As FoundFiles  (read-only)`
- `PropertyTests As PropertyTests  (read-only)`
- `SearchScopes As SearchScopes  (read-only)`
- `SearchFolders As SearchFolders  (read-only)`
- `FileTypes As FileTypes  (read-only)`

## Methods (3)

- `Execute([SortBy As MsoSortBy], [SortOrder As MsoSortOrder], [AlwaysAccurate As Boolean]) As Long`
- `NewSearch()`
- `RefreshScopes()`
