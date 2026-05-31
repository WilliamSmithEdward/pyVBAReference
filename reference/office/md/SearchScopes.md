# SearchScopes

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0366-0000-0000-C000-000000000046}  

A collection of SearchScope objects.

**Remarks:** You can't add or remove SearchScope objects from the SearchScopes collection.

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SearchScopes object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SearchScopes object was created. Read-only.
- `Item As SearchScope  (read-only)`  
  Gets a SearchScope object that corresponds to an area in which to perform a file search, such as local drives or Microsoft Outlook folders. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the SearchScopes object. Read-only.
- `_NewEnum As IUnknown  (read-only)`
