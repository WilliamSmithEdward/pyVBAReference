# IBlogExtensibility

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C4-0000-0000-C000-000000000046}  

An object that provides the ability to manipulate blog entries.

## Methods (8)

- `BlogProviderProperties(BlogProvider As String, FriendlyName As String, CategorySupport As MsoBlogCategorySupport, Padding As Boolean)`  
  Contains information about the provider.
    - `BlogProvider As String` (required): The name of the blog provider.
    - `FriendlyName As String` (required): Represents the name displayed in the user interface.
    - `CategorySupport As MsoBlogCategorySupport` (required): Represents how many categories are supported by the provider.
    - `Padding As Boolean` (required): Specifies whether table padding is recognized.
- `SetupBlogAccount(Account As String, ParentWindow As Long, Document As Object, NewAccount As Boolean, ShowPictureUI As Boolean)`  
  Called from the Choose Account dialog when the provider's name is chosen in the Blog Host drop-down, or when the user requests to change a provider's account in the Blog Accounts dialog box.
    - `Account As String` (required): Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.
    - `ParentWindow As Long` (required): Contains the HWND for the window that Microsoft Word is calling from.
    - `Document As Object` (required): The current document.
    - `NewAccount As Boolean` (required): Indicates whether this is a new account.
    - `ShowPictureUI As Boolean` (required): Indicates whether Word's picture user interface needs to be displayed.
- `GetUserBlogs(Account As String, ParentWindow As Long, Document As Object, BlogNames As SAFEARRAY(String), BlogIDs As SAFEARRAY(String), BlogURLs As SAFEARRAY(String))`  
  Returns the list and details of user blogs associated with the specified account.
    - `Account As String` (required): Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.
    - `ParentWindow As Long` (required): Contains the HWND for the window that Microsoft Word is calling from.
    - `Document As Object` (required): The current document.
- `GetRecentPosts(Account As String, ParentWindow As Long, Document As Object, PostTitles As SAFEARRAY(String), PostDates As SAFEARRAY(String), PostIDs As SAFEARRAY(String))`  
  Returns the list of the user's last fifteen blog posts that Microsoft Word then displays in the Open Existing Post dialog. This method does not actually return the blog post contents.
    - `Account As String` (required): Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.
    - `ParentWindow As Long` (required): Contains the HWND for the window that Microsoft Word is calling from.
    - `Document As Object` (required): The current document.
- `Open(Account As String, PostID As String, ParentWindow As Long, xHTML As String, Title As String, DatePosted As String, Categories As SAFEARRAY(String))`  
  Opens the blog specified by the blog ID. It's called by the Open Existing Post dialog based on the item selected by the user.
    - `Account As String` (required): Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.
    - `PostID As String` (required): The ID of the post.
    - `ParentWindow As Long` (required): Contains the HWND for the window that Microsoft Word is calling from.
    - `xHTML As String` (required): Represents the xHTML of the current document.
    - `Title As String` (required): The title of the post.
    - `DatePosted As String` (required): The date the entry was posted.
- `PublishPost(Account As String, ParentWindow As Long, Document As Object, xHTML As String, Title As String, DateTime As String, Categories As SAFEARRAY(String), Draft As Boolean, PostID As String, PublishMessage As String)`  
  Hands off the current post so that it can be published by the provider.
    - `Account As String` (required): Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.
    - `ParentWindow As Long` (required): Contains the HWND for the window that Microsoft Word is calling from.
    - `Document As Object` (required): The current document.
    - `xHTML As String` (required): Represents the xHTML of the current document.
    - `Title As String` (required): The title of the post.
    - `DateTime As String` (required): The date the entry was posted.
    - `Draft As Boolean` (required): Specifies whether this is a draft version of the post.
    - `PostID As String` (required): The ID of the original post if this post has been republished.
    - `PublishMessage As String` (required): Specifies what is displayed in the publish bar.
- `RepublishPost(Account As String, ParentWindow As Long, Document As Object, PostID As String, xHTML As String, Title As String, DateTime As String, Categories As SAFEARRAY(String), Draft As Boolean, PublishMessage As String)`  
  Hands off the current post so it can be republished by the provider.
    - `Account As String` (required): Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.
    - `ParentWindow As Long` (required): Contains the HWND for the window that Microsoft Word is calling from.
    - `Document As Object` (required): The current document.
    - `PostID As String` (required): The ID of the original post.
    - `xHTML As String` (required): Represents the xHTML of the current document.
    - `Title As String` (required): The title of the post.
    - `DateTime As String` (required): The date the entry was posted.
    - `Draft As Boolean` (required): Specifies whether this is a draft version of the post.
    - `PublishMessage As String` (required): Specifies what is displayed in the publish bar.
- `GetCategories(Account As String, ParentWindow As Long, Document As Object, Categories As SAFEARRAY(String))`  
  This method returns the list of blog categories for an account so that Microsoft Word can populate the categories drop-down list.
    - `Account As String` (required): Represents the GUID of the account registry key.
    - `ParentWindow As Long` (required): Represents the HWND of the host window.
    - `Document As Object` (required): The current document.
