# LanguageSettings

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0353-0000-0000-C000-000000000046}  

Returns information about the language settings in a Microsoft Office application.

**Remarks:** Use Application.LanguageSettings.LanguageID(_MsoAppLanguageID_), where MsoAppLanguageID is a constant used to return locale identifier (LCID) information to the specified application.

**Example:**

```vba
MsgBox "The following locale IDs are registered " & _
 "for this application: Install Language - " & _
 Application.LanguageSettings.LanguageID(msoLanguageIDInstall) & _
 " User Interface Language - " & _
 Application.LanguageSettings.LanguageID(msoLanguageIDUI) & _
 " Help Language - " & _
 Application.LanguageSettings.LanguageID(msoLanguageIDHelp)
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the LanguageSettings object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the LanguageSettings object was created. Read-only.
- `LanguageID As Long  (read-only)`  
  Gets an MsoAppLanguageID constant representing the locale identifier (LCID) for the install language, the user interface language, or the Help language. Read-only.
- `LanguagePreferredForEditing As Boolean  (read-only)`  
  Gets True if the value for the MsoLanguageID constant has been identified in the Windows registry as a preferred language for editing. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ODSOFilter object. Read-only.
