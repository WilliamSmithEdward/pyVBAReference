# Printers

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {DBC51760-A8ED-11D3-A0DD-00C04F68712B}  

The Printers collection contains Printer objects representing all the printers available on the current system.

**Remarks:** Use the Printers property of the Application object to return the Printers collection. You can enumerate through the Printers collection by using the For Each...Next statement. Refer to an individual Printer object in the Printers collection either by referring to the printer by name, or by referring to its index within the collection. The Printers collection is indexed beginning with zero. If you refer to a printer by its index, the first printer is Printers(0), the second printer is Printers(1), and so on. You can't add or delete a Printer object from the Printers collection.

**Example:**

```vba
Dim prtLoop As Printer

For Each prtLoop In Application.Printers
 With prtLoop
 MsgBox "Device name: " & .DeviceName & vbCr _
 & "Driver name: " & .DriverName & vbCr _
 & "Port: " & .Port
 End With
Next prtLoop
```

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As _Printer  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only Printer.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
