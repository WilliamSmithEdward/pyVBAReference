# OLEFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024441-0000-0000-C000-000000000046}  

Contains OLE object properties.

**Remarks:** If the Shape object doesn't represent a linked or embedded object, the OLEFormat property of the Shape object fails.

**Example:**

```vba
Worksheets(1).Shapes(1).OLEFormat.Activate
```

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Object As Object  (read-only)`  
  Returns the OLE Automation object associated with this OLE object. Read-only Object.
- `progID As String  (read-only)`  
  Returns the programmatic identifiers for the object. Read-only String.

## Methods (2)

- `Activate()`  
  Activates the current OLE object.
- `Verb([Verb As Variant])`  
  Sends a verb to the server of the specified OLE object.
    - `Verb As Variant` (optional): The verb that the server of the OLE object should act on. If this argument is omitted, the default verb is sent. The available verbs are determined by the object's source application. Typical verbs for an OLE object are Open and Primary (represented by the XlOLEVerb constants xlOpen and xlPrimary).
