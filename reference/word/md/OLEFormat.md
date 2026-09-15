# OLEFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020933-0000-0000-C000-000000000046}  

Represents the OLE characteristics (other than linking) for an OLE object, ActiveX control, or field.

**Remarks:** Use the OLEFormat property for a shape, inline shape, or field to return the OLEFormat object. The following example displays the class type for the first shape on the active document. Not all types of shapes, inline shapes, and fields have OLE capabilities. Use the Type property for the Shape and InlineShape objects to determine what category the specified shape or inline shape falls into. The Type property for a Field object returns the type of field. Use the Activate, Edit, Open, and DoVerb methods to automate an OLE object. Use the Object property to return an object that represents an ActiveX control or OLE object. With this object, you can use the properties and methods of the container application or the ActiveX control.

## Properties (13)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OLEFormat object.
- `ClassType As String  (read/write)`  
  Returns or sets the class type for the specified OLE object, picture, or field. Read/write String.
- `DisplayAsIcon As Boolean  (read/write)`  
  True if the specified object is displayed as an icon. Read/write Boolean.
- `IconName As String  (read/write)`  
  Returns or sets the program file in which the icon for an OLE object is stored. Read/write String.
- `IconPath As String  (read-only)`  
  Returns the path of the file in which the icon for an OLE object is stored. Read-only String.
- `IconIndex As Long  (read/write)`  
  Returns or sets the icon that is used when the DisplayAsIcon property is True. Read/write Long.
- `IconLabel As String  (read/write)`  
  Returns or sets the text displayed below the icon for an OLE object. Read/write String.
- `Label As String  (read-only)`  
  Returns a string that's used to identify the portion of the source file that's being linked. Read-only String.
- `Object As Object  (read-only)`  
  Returns an Object that represents the specified OLE object's top-level interface. .
- `ProgID As String  (read-only)`  
  Returns the programmatic identifier (ProgID) for the specified OLE object. Read-only String.
- `PreserveFormattingOnUpdate As Boolean  (read/write)`  
  True preserves formatting done in Microsoft Word to a linked OLE object, such as a table linked to a Microsoft Excel spreadsheet. Read/write Boolean.

## Methods (6)

- `Activate()`  
  Activates the specified OLEFormat object.
- `Edit()`  
  Opens the specified OLE object for editing in the application it was created in.
- `Open()`  
  Opens the specified OLEFormat object.
- `DoVerb([VerbIndex As Variant])`  
  Requests that an OLE object perform one of its available verbs&mdash;the actions an OLE object takes to activate its contents.
    - `VerbIndex As Variant` (optional): The verb that the OLE object should perform. If this argument is omitted, the default verb is sent. If the OLE object does not support the requested verb, an error will occur. Can be any WdOLEVerb constant.
- `ConvertTo([ClassType As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant])`  
  Converts the specified OLE object from one class to another, making it possible for you to edit the object in a different server application or change how the object is displayed in the document.
    - `ClassType As Variant` (optional): The name of the application used to activate the OLE object. You can see a list of the available applications in the Object type box on the Create New tab in the Object dialog box. You can find the ClassType string by inserting an object as an inline shape and then viewing the field codes. The class type of the object follows either the word "EMBED" or the word "LINK."
    - `DisplayAsIcon As Variant` (optional): True to display the OLE object as an icon. The default value is False.
    - `IconFileName As Variant` (optional): The file that contains the icon to be displayed.
    - `IconIndex As Variant` (optional): The index number of the icon within IconFileName. The order of icons in the specified file corresponds to the order in which the icons appear in the Change Icon dialog box (Insert Object dialog box) when the Display as icon check box is selected. The first icon in the file has the index number 0 (zero). If an icon with the given index number doesn't exist in IconFileName, the icon with the index number 1 (the second icon in the file) is used. The default value is 0 (zero).
    - `IconLabel As Variant` (optional): A label (caption) to be displayed beneath the icon.
- `ActivateAs(ClassType As String)`  
  Sets the Windows registry value that determines the default application used to activate the specified OLE object.
    - `ClassType As String` (required): The name of the application in which an OLE object is opened. To see a list of object types that the OLE object can be activated as, click the object and then open the Convert dialog box. You can find the ClassType string by inserting an object as an inline shape and then viewing the field codes. The class type of the object follows either the word "EMBED" or the word "LINK."
