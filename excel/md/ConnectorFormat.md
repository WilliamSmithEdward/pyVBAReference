# ConnectorFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002443E-0000-0000-C000-000000000046}  

Contains properties and methods that apply to connectors.

**Remarks:** A connector is a line that attaches two other shapes at points called connection sites. If you rearrange shapes that are connected, the geometry of the connector will be automatically adjusted so that the shapes remain connected. Connection sites are generally numbered according to the rules presented in the following table. Use the ConnectorFormat property of the Shape object to return a ConnectorFormat object. Use the BeginConnect and EndConnect methods to attach the ends of the connector to other shapes in the document. Use the RerouteConnections method of the Shape object to automatically find the shortest path between the two shapes connected by the connector. Use the Connector property to see whether a shape is a connector. >

**Example:**

```vba
Set mainshape = ActiveWindow.Selection.ShapeRange(1)
With mainshape
 bx = .Left + .Width + 50
 by = .Top + .Height + 50
End With
With ActiveSheet
 For j = 1 To mainshape.ConnectionSiteCount
 With .Shapes.AddConnector(msoConnectorStraight, _
 bx, by, bx + 50, by + 50)
 .ConnectorFormat.EndConnect mainshape, j
 .ConnectorFormat.Type = msoConnectorElbow
 .Line.ForeColor.RGB = RGB(255, 0, 0)
 l = .Left
 t = .Top
 End With
 With .Shapes.AddTextbox(msoTextOrientationHorizontal, _
 l, t, 36, 14)
 .Fill.Visible = False
 .Line.Visible = False
 .TextFrame.Characters.Text = j
 End With
 Next j
End With
```

## Properties (10)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `BeginConnected As MsoTriState  (read-only)`  
  True if the beginning of the specified connector is connected to a shape. Read-only MsoTriState.
- `BeginConnectedShape As Shape  (read-only)`  
  Returns a Shape object that represents the shape that the beginning of the specified connector is attached to. Read-only.
- `BeginConnectionSite As Long  (read-only)`  
  Returns an integer that specifies the connection site that the beginning of a connector is connected to. Read-only Long.
- `EndConnected As MsoTriState  (read-only)`  
  msoTrue if the end of the specified connector is connected to a shape. Read-only MsoTriState.
- `EndConnectedShape As Shape  (read-only)`  
  Returns a Shape object that represents the shape that the end of the specified connector is attached to. Read-only.
- `EndConnectionSite As Long  (read-only)`  
  Returns an integer that specifies the connection site that the end of a connector is connected to. Read-only Long.
- `Type As MsoConnectorType  (read/write)`  
  Returns or sets an MsoConnectorType value that represents the connector format type.

## Methods (4)

- `BeginConnect(ConnectedShape As Shape, ConnectionSite As Long)`  
  Attaches the beginning of the specified connector to a specified shape. If there's already a connection between the beginning of the connector and another shape, that connection is broken.
    - `ConnectedShape As Shape` (required): The shape to attach the beginning of the connector to. The specified Shape object must be in the same Shapes collection as the connector.
    - `ConnectionSite As Long` (required): A connection site on the shape specified by _ConnectedShape_. Must be an integer between 1 and the integer returned by the ConnectionSiteCount property of the specified shape. If you want the connector to automatically find the shortest path between the two shapes it connects, specify any valid integer for this argument, and then use the RerouteConnections method after the connector is attached to shapes at both ends.
- `BeginDisconnect()`  
  Detaches the beginning of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector; the beginning of the connector remains positioned at a connection site but is no longer connected.
- `EndConnect(ConnectedShape As Shape, ConnectionSite As Long)`  
  Attaches the end of the specified connector to a specified shape. If there's already a connection between the end of the connector and another shape, that connection is broken.
    - `ConnectedShape As Shape` (required): The shape to attach the end of the connector to. The specified Shape object must be in the same Shapes collection as the connector.
    - `ConnectionSite As Long` (required): Must be an integer between 1 and the integer returned by the ConnectionSiteCount property of the specified shape. If you want the connector to automatically find the shortest path between the two shapes it connects, specify any valid integer for this argument and then use the RerouteConnections method after the connector is attached to shapes at both ends.
- `EndDisconnect()`  
  Detaches the end of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector; the end of the connector remains positioned at a connection site but is no longer connected.
