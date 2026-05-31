# IAssistance

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {4291224C-DEFE-485B-8E69-6CF8AA85CB76}  

Provides a means for developers to create a customized help experience for users within Microsoft Office.

**Remarks:** The Assistance property returns an IAssistance object. The IAssistance object exposes methods that allow developers to display help topics in the Office Help Viewer or to display help topics that ship with Office in the Help window of the host application. Developers either pass specific Help IDs to the help system or pass specific search queries. Help IDs have to be explicitly added to the Help file in order for the Help ID to return the help topic.

**Example:**

```vba
Sub DisplayHelpTopic()
 Application.Assistance.ShowHelp "xlmain11.chm60407", ""
 Application.Assistance.ShowHelp "vbaxl10.chm65879", "DEV"
End Sub
```

## Methods (4)

- `ShowHelp([HelpId As String], [Scope As String])`  
  Displays the help topic specified by its ID in the Office Help Viewer, or for help topics that ship with Office, in the Help window of the current Office application.
    - `HelpId As String` (optional): The ID of the help topic.
    - `Scope As String` (optional): The namespace registered within the host application.
- `SearchHelp(Query As String, [Scope As String])`  
  Performs a search from the Office Help Viewer based on one or more keywords. Keywords can be a word or a phrase.
    - `Query As String` (required): Represents the search keyword or phrase.
    - `Scope As String` (optional): The namespace registered within the host application.
- `SetDefaultContext(HelpId As String)`  
  Sets a help topic as the default topic that will be displayed when the user opens a help window.
    - `HelpId As String` (required): The ID of the default help topic.
- `ClearDefaultContext(HelpId As String)`  
  Clears the default help topic previously defined in the SetDefaultContext method.
    - `HelpId As String` (required): The ID of the default help topic.
