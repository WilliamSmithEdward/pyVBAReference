# MacroError

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3F1B1773-65CB-4DB9-9FC6-ACED47DB285A}  

Represents the properties of a run-time error that occurs in a macro.

**Remarks:** When an error occurs in a macro, information about the error is stored in the MacroError object. If you have not used the OnError action to suppress error messages, the macro stops and the error information is displayed in a standard error message. However, if you have used the OnError action to suppress error messages, you may want to use the information stored in the MacroError object in a condition or a custom error message. After an error has been handled, the information in the MacroError object is out of date, so it is a good idea to clear the object by using the ClearMacroError action. This resets the error number in the MacroError object back to zero, and clears any other information about the error that is stored in the object, such as the error description, macro name, action name, condition, and arguments. This way, you can inspect the MacroError object again later to see if another error has occurred. The MacroError object contains information about only one error at a time. If more than one error has occurred in a macro, the MacroError object contains information about only the last one.

## Properties (6)

- `Condition As String  (read-only)`  
  Gets the condition of the macro action that was executing when an error occurred. Read-only String.
- `ActionName As String  (read-only)`  
  Gets the name of the macro action that was executing when an error occurred. Read-only.
- `Arguments As String  (read-only)`  
  Gets the arguments specified for the macro action that was executing when an error occurred. Read-only String.
- `Description As String  (read-only)`  
  Gets the text that represents the current error message. Read-only String.
- `Number As Long  (read-only)`  
  Gets the current error number. Read-only Long.
- `MacroName As String  (read-only)`  
  Gets the name of the macro that was running when an error occurred. Read-only String.
