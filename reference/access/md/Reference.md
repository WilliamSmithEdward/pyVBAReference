# Reference

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {EB106212-9C89-11CF-A2B3-00A0C90542FF}  

The Reference object refers to a reference set to another application's or project's type library.

**Remarks:** When you create a Reference object, you set a reference dynamically from Visual Basic. The Reference object is a member of the References collection. To refer to a particular Reference object in the References collection, use any of the following syntax forms. The following example refers to the Reference object that represents the reference to the Microsoft Access type library.

## Properties (9)

- `Collection As References  (read-only)`  
  The Collection property returns a reference to the collection that contains an object. Read-only References object.
- `Name As String  (read-only)`  
  Use the Name property to determine the string expression that identifies the name of an object. Read-only String.
- `Guid As String  (read-only)`  
  The GUID property of a Reference object returns a GUID that identifies a type library in the Windows Registry. Read-only String.
- `Major As Long  (read-only)`  
  The Major property of a Reference object returns a read-only Long value indicating the major version number of an application to which you have set a reference.
- `Minor As Long  (read-only)`  
  The Minor property of a Reference object returns a Long value indicating the minor version number of the application to which you have set a reference.
- `FullPath As String  (read-only)`  
  The FullPath property returns a string containing the path and file name of the referenced type library.
- `BuiltIn As Boolean  (read-only)`  
  The BuiltIn property returns a Boolean value indicating whether a Reference object points to a default reference that's necessary for Microsoft Access to function properly. Read-only Boolean.
- `IsBroken As Boolean  (read-only)`  
  The IsBroken property returns a Boolean value indicating whether a Reference object points to a valid reference in the Windows Registry. Read-only Boolean.
- `Kind As vbext_RefKind  (read-only)`  
  The Kind property indicates the type of reference that a Reference object represents. Read-only vbext_RefKind.
