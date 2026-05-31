# IFind

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0337-0000-0000-C000-000000000046}  

## Properties (22)

- `SearchPath As String  (read/write)`
- `Name As String  (read/write)`
- `SubDir As Boolean  (read/write)`
- `Title As String  (read/write)`
- `Author As String  (read/write)`
- `Keywords As String  (read/write)`
- `Subject As String  (read/write)`
- `Options As MsoFileFindOptions  (read/write)`
- `MatchCase As Boolean  (read/write)`
- `Text As String  (read/write)`
- `PatternMatch As Boolean  (read/write)`
- `DateSavedFrom As Variant  (read/write)`
- `DateSavedTo As Variant  (read/write)`
- `SavedBy As String  (read/write)`
- `DateCreatedFrom As Variant  (read/write)`
- `DateCreatedTo As Variant  (read/write)`
- `View As MsoFileFindView  (read/write)`
- `SortBy As MsoFileFindSortBy  (read/write)`
- `ListBy As MsoFileFindListBy  (read/write)`
- `SelectedFile As Long  (read/write)`
- `Results As IFoundFiles  (read-only)`
- `FileType As Long  (read/write)`

## Methods (5)

- `Show() As Long`
- `Execute()`
- `Load(bstrQueryName As String)`
- `Save(bstrQueryName As String)`
- `Delete(bstrQueryName As String)`
