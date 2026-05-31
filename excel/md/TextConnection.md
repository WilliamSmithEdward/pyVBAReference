# TextConnection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244D3-0000-0000-C000-000000000046}  

Contains Service Contract settings that enable Microsoft Excel to connect to a Data Feed data source.

## Properties (22)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TextConnection object. Read-only.
- `Connection As Variant  (read/write)`  
  Returns or sets a string that contains text file names that enable Microsoft Excel to connect to text data sources. Read/write Variant.
- `TextFileHeaderRow As Boolean  (read/write)`  
  Returns or sets a value that specifies whether the first row (from the starting row) should be treated as a header row. Read/write Boolean.
- `TextFileColumnDataTypes As Variant  (read/write)`  
  Returns or sets an ordered array of constants that specify the data types applied to the corresponding columns in the text file that you are importing into a query table. The default constant for each column is xlGeneral. Read/write Variant.
- `TextFileCommaDelimiter As Boolean  (read/write)`  
  True if the comma is the delimiter when you import a text file into a query table. False if you want to use some other character as the delimiter. The default value is False. Read/write Boolean.
- `TextFileConsecutiveDelimiter As Boolean  (read/write)`  
  True if consecutive delimiters are treated as a single delimiter when you import a text file into a query table. The default value is False. Read/write Boolean.
- `TextFileDecimalSeparator As String  (read/write)`  
  Returns or sets the decimal separator character that Microsoft Excel uses when you import a text file into a query table. The default is the system decimal separator character. Read/write String.
- `TextFileFixedColumnWidths As Variant  (read/write)`  
  Returns or sets an array of integers that correspond to the widths of the columns (in characters) in the text file that you are importing into a query table. Valid widths are from 1 through 32767 characters. Read/write Variant.
- `TextFileOtherDelimiter As String  (read/write)`  
  Returns or sets the character used as the delimiter when you import a text file into a query table. The default value is null. Read/write String.
- `TextFileParseType As XlTextParsingType  (read/write)`  
  Returns or sets the column format for the data in the text file that you are importing into a query table. Read/write XlTextParsingType enumeration.
- `TextFilePlatform As XlPlatform  (read/write)`  
  Returns or sets the origin of the text file that you are importing into the query table. This property determines which code page is used during the data import. Read/write XlPlatform.
- `TextFilePromptOnRefresh As Boolean  (read/write)`  
  True if you want to specify the name of the imported text file each time the query table is refreshed. The Import Text File dialog box allows you to specify the path and file name. The default value is False. Read/write Boolean.
- `TextFileSemicolonDelimiter As Boolean  (read/write)`  
  True if the semicolon is the delimiter when you import a text file into a query table, and if the value of the TextFileParseType property is xlDelimited. The default value is False. Read/write Boolean.
- `TextFileSpaceDelimiter As Boolean  (read/write)`  
  True if the space character is the delimiter when you import a text file into a query table. The default value is False. Read/write Boolean.
- `TextFileStartRow As Long  (read/write)`  
  Returns or sets the row number at which text parsing will begin when you import a text file into a query table. Valid values are integers from 1 through 32767. The default value is 1. Read/write Long.
- `TextFileTabDelimiter As Boolean  (read/write)`  
  True if the tab character is the delimiter when you import a text file into a query table. The default value is False. Read/write Boolean.
- `TextFileTextQualifier As XlTextQualifier  (read/write)`  
  Returns or sets the text qualifier when you import a text file into a query table. The text qualifier specifies that the enclosed data is in text format. Read/write XlTextQualifier.
- `TextFileThousandsSeparator As String  (read/write)`  
  Returns or sets the thousands separator character that Microsoft Excel uses when you import a text file into a query table. The default is the system thousands separator character. Read/write String.
- `TextFileTrailingMinusNumbers As Boolean  (read/write)`  
  True for Microsoft Excel to treat numbers imported as text that begin with a - (minus) symbol as a negative symbol. False for Excel to treat numbers imported as text that begin with a - symbol as text. Read/write Boolean.
- `TextFileVisualLayout As XlTextVisualLayoutType  (read/write)`  
  Returns or sets an XlTextVisualLayoutType value that indicates whether the visual layout of the text being imported is left-to-right or right-to-left. Read/write.
