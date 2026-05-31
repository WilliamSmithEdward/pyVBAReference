# IBlogPictureExtensibility

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03C5-0000-0000-C000-000000000046}  

Represents an object that provides the ability to manipulate blog images.

**Remarks:** Pictures are stored with picture providers in an account set up by the user.

## Methods (3)

- `BlogPictureProviderProperties(BlogPictureProvider As String, FriendlyName As String)`  
  Enables picture providers to offer themselves as an upload location for blog pictures.
    - `BlogPictureProvider As String` (required): The ID of the picture provider.
    - `FriendlyName As String` (required): The friendly name of the picture provider.
- `CreatePictureAccount(Account As String, BlogProvider As String, ParentWindow As Long, Document As Object)`  
  Allows a picture provider to display the user interface needed to guide the user through setting up a picture account.
    - `Account As String` (required): Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.
    - `BlogProvider As String` (required): The ID of the provider.
    - `ParentWindow As Long` (required): Contains the HWND for the window that Microsoft Word is calling from.
    - `Document As Object` (required): The current document.
- `PublishPicture(Account As String, ParentWindow As Long, Document As Object, Image As IUnknown, PictureURI As String, ImageType As Long)`  
  Used to post a picture object to its final destination in a blog.
    - `Account As String` (required): Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.
    - `ParentWindow As Long` (required): Contains the HWND for the window that Microsoft Word is calling from.
    - `Document As Object` (required): The current document.
    - `Image As IUnknown` (required): Represents the name of the image file.
    - `PictureURI As String` (required): The URI of the picture.
