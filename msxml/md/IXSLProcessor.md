# IXSLProcessor

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {2933BF92-7B36-11D2-B20E-00C04F983E60}  

IXSLProcessor Interface

## Properties (7)

- `input As Variant  (read/write)`  
  XML input tree to transform
- `ownerTemplate As IXSLTemplate  (read-only)`  
  template object used to create this processor object
- `startMode As String  (read-only)`  
  starting XSL mode
- `startModeURI As String  (read-only)`  
  namespace of starting XSL mode
- `output As Variant  (read/write)`  
  custom stream object for transform output
- `readyState As Long  (read-only)`  
  current state of the processor
- `stylesheet As IXMLDOMNode  (read-only)`  
  current stylesheet being used

## Methods (5)

- `setStartMode(mode As String, [namespaceURI As String])`  
  set XSL mode and it's namespace
- `transform() As Boolean`  
  start/resume the XSL transformation process
- `reset()`  
  reset state of processor and abort current transform
- `addParameter(baseName As String, parameter As Variant, [namespaceURI As String])`  
  set <xsl:param> values
- `addObject(obj As Object, namespaceURI As String)`  
  pass object to stylesheet
