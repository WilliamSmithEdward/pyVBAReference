# WebServices

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {371BF535-7914-4474-BAE9-27281D431237}  

Represents the collection of Data Services data connections installed in the database.

**Remarks:** Use the WebServices property of the Application object to return the collection of installed Data Services data connections. Use the following steps to install a Data Services data connection in your database: 1. Obtain a Data Services data connection file of the data source that you want to connect to. 2. On the ribbon, choose the External Data tab. 3. In the Import & Link group, choose the More drop-down, and then choose Data Services. 4. At the bottom of the Create Link to Data Services dialog box, choose Install New Connection. 5. In the Select a connection definition file dialog box, browse to and select the XML file that contains the description of the Data Service.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As WebService  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only Object.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
