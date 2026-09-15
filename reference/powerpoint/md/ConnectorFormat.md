# ConnectorFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493481-5A91-11CF-8700-00AA0060263B}  

Contains properties and methods that apply to connectors.

**Remarks:** A connector is a line that attaches two other shapes at points called connection sites. If you rearrange shapes that are connected, the geometry of the connector will be automatically adjusted so that the shapes remain connected. Use the ConnectorFormat property to return a ConnectorFormat object. Use the BeginConnect and EndConnect methods to attach the ends of the connector to other shapes in the document. Use the RerouteConnections method to automatically find the shortest path between the two shapes connected by the connector. Use the Connector property to see whether a shape is a connector. Note that you assign a size and a position when you add a connector to the Shapes collection, but the size and position are automatically adjusted when you attach the beginning and end of the connector to other shapes in the collection. Therefore, if you intend to attach a connector to other shapes, the initial size and position you specify are irrelevant.

**Example:**

```vba
Set myDocument = ActivePresentation.Slides(1)

Set s = myDocument.Shapes

Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)

Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)

With s.AddConnector(msoConnectorCurve, 0, 0, 0, 0).ConnectorFormat

    .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1

    .EndConnect ConnectedShape:=secondRect, ConnectionSite:=1

    .Parent.RerouteConnections

End With
```

## Properties (10)

- `Application As Object  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `BeginConnected As MsoTriState  (read-only)`  
  Determines whether the beginning of the specified connector is connected to a shape. Read/write.
- `BeginConnectedShape As Shape  (read-only)`  
  Returns a Shape object that represents the shape that the beginning of the specified connector is attached to. Read-only.
- `BeginConnectionSite As Long  (read-only)`  
  Returns an integer that specifies the connection site that the beginning of a connector is connected to. Read-only.
- `EndConnected As MsoTriState  (read-only)`  
  Determines whether the end of the specified connector is connected to a shape. Read-only.
- `EndConnectedShape As Shape  (read-only)`  
  Returns a Shape object that represents the shape that the end of the specified connector is attached to. Read-only.
- `EndConnectionSite As Long  (read-only)`  
  Returns an integer that specifies the connection site that the end of a connector is connected to. Read-only.
- `Type As MsoConnectorType  (read/write)`  
  Represents the type of connector. Read/write.

## Methods (4)

- `BeginConnect(ConnectedShape As Shape, ConnectionSite As Long)`  
  Attaches the beginning of the specified connector to a specified shape.
    - `ConnectedShape As Shape` (required): The shape to attach the beginning of the connector to. The specified Shape object must be in the same Shapes collection as the connector.
    - `ConnectionSite As Long` (required): A connection site on the shape specified by ConnectedShape. Must be an integer between 1 and the integer returned by the ConnectionSiteCount property of the specified shape. If you want the connector to automatically find the shortest path between the two shapes it connects, specify any valid integer for this argument and then use the RerouteConnections method after the connector is attached to shapes at both ends.
- `BeginDisconnect()`  
  Detaches the beginning of the specified connector from the shape it is attached to.
- `EndConnect(ConnectedShape As Shape, ConnectionSite As Long)`  
  Attaches the end of the specified connector to a specified shape.
    - `ConnectedShape As Shape` (required): The shape to attach the end of the connector to. The specified Shape object must be in the same Shapes collection as the connector.
    - `ConnectionSite As Long` (required): A connection site on the shape specified by ConnectedShape. Must be an integer between 1 and the integer returned by the ConnectionSiteCount property of the specified shape. If you want the connector to automatically find the shortest path between the two shapes it connects, specify any valid integer for this argument and then use the RerouteConnections method after the connector is attached to shapes at both ends.
- `EndDisconnect()`  
  Detaches the end of the specified connector from the shape it is attached to. This method doesn't alter the size or position of the connector: the end of the connector remains positioned at a connection site but is no longer connected. Use the BeginDisconnect method to detach the beginning of the connector from a shape.
