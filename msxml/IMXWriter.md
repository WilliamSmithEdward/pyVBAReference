# IMXWriter

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {4D7FF4BA-1565-4EA8-94E1-6E724A46F98D}  

IMXWriter interface

## Properties (8)

- `output As Variant  (read/write)`  
  Set or get the output.
- `encoding As String  (read/write)`  
  Set or get the output encoding.
- `byteOrderMark As Boolean  (read/write)`  
  Determine whether or not to write the byte order mark
- `indent As Boolean  (read/write)`  
  Enable or disable auto indent mode.
- `standalone As Boolean  (read/write)`  
  Set or get the standalone document declaration.
- `omitXMLDeclaration As Boolean  (read/write)`  
  Determine whether or not to omit the XML declaration.
- `version As String  (read/write)`  
  Set or get the xml version info.
- `disableOutputEscaping As Boolean  (read/write)`  
  When enabled, the writer no longer escapes out its input when writing it out.

## Methods (1)

- `flush()`  
  Flushes all writer buffers forcing the writer to write to the underlying output object
