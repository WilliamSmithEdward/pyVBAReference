# MultiThreadedCalculation

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244B1-0000-0000-C000-000000000046}  

Returns or sets the concurrent calculation mode.

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Enabled As Boolean  (read/write)`  
  The Enabled property allows MultiThreadedCalculation objects to be enabled or disabled at run time. Read/write.
- `ThreadMode As XlThreadMode  (read/write)`  
  Returns or sets the thread mode for the specified MultiThreadedCalculation object. Read/write XlThreadMode.
- `ThreadCount As Long  (read/write)`  
  Gets the total count of the process threads that are a part of the specified MultiThreadedCalculation object.
