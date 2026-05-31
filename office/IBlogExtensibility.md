# IBlogExtensibility

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C4-0000-0000-C000-000000000046}  

## Methods (8)

- `BlogProviderProperties(BlogProvider As String, FriendlyName As String, CategorySupport As MsoBlogCategorySupport, Padding As Boolean)`
- `SetupBlogAccount(Account As String, ParentWindow As Long, Document As Object, NewAccount As Boolean, ShowPictureUI As Boolean)`
- `GetUserBlogs(Account As String, ParentWindow As Long, Document As Object, BlogNames As SAFEARRAY(String), BlogIDs As SAFEARRAY(String), BlogURLs As SAFEARRAY(String))`
- `GetRecentPosts(Account As String, ParentWindow As Long, Document As Object, PostTitles As SAFEARRAY(String), PostDates As SAFEARRAY(String), PostIDs As SAFEARRAY(String))`
- `Open(Account As String, PostID As String, ParentWindow As Long, xHTML As String, Title As String, DatePosted As String, Categories As SAFEARRAY(String))`
- `PublishPost(Account As String, ParentWindow As Long, Document As Object, xHTML As String, Title As String, DateTime As String, Categories As SAFEARRAY(String), Draft As Boolean, PostID As String, PublishMessage As String)`
- `RepublishPost(Account As String, ParentWindow As Long, Document As Object, PostID As String, xHTML As String, Title As String, DateTime As String, Categories As SAFEARRAY(String), Draft As Boolean, PublishMessage As String)`
- `GetCategories(Account As String, ParentWindow As Long, Document As Object, Categories As SAFEARRAY(String))`
