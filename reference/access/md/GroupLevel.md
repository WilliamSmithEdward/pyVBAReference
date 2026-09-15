# GroupLevel

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {BC9E4356-F037-11CD-8701-00AA003F0F07}  

Use the GroupLevel property in Visual Basic to refer to the group level that you are grouping or sorting on in a report.

**Remarks:** The GroupLevel property setting is an array in which each entry identifies a group level. To refer to a group level, use this syntax: GroupLevel (n) The number n is the group level, starting with 0. The first field or expression that you group on is group level 0, the second is group level 1, and so on. You can have up to 10 group levels (0 to 9). The following sample settings show how you use the GroupLevel property to refer to a group level. Use this property only by using Visual Basic to set the SortOrder, GroupOn, GroupInterval, KeepTogether, and ControlSource properties. You set these properties in the Open event procedure of a report. In reports, you can group or sort on more than one field or expression. Each field or expression that you group or sort on is a group level. You specify the fields and expressions to sort and group on by using the CreateGroupLevel method. If a group is already defined for a report (the GroupLevel property is set to 0), you can use the ControlSource property to change the group level in the report's Open event procedure.

## Properties (10)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Properties As Properties  (read-only)`  
  Returns a reference to a control's Properties collection object. Read-only.
- `ControlSource As String  (read/write)`  
  Use the ControlSource property to specify what data appears in a control. You can display and edit data bound to a field in a table, query, or SQL statement. You can also display the result of an expression. Read/write String.
- `SortOrder As Boolean  (read/write)`  
  You use the SortOrder property to specify the sort order for fields and expressions in a report. For example, if you are printing a list of suppliers, you can sort the records alphabetically by company name. Read/write Boolean.
- `GroupHeader As Boolean  (read/write)`  
  Use the GroupHeader property to create a group header for a selected field or expression in a report. Read/write Boolean.
- `GroupFooter As Boolean  (read/write)`  
  Use the GroupFooter property to create a group footer for a selected field or expression in a report. Read/write Boolean.
- `GroupOn As Integer  (read/write)`  
  Use the GroupOn property in a report to specify how to group data in a field or expression by data type. For example, this property lets you group a Date field by month. Read/write Integer.
- `GroupInterval As Long  (read/write)`  
  Use the GroupInterval property with the GroupOn property to specify how records are grouped in a report. Read/write Long.
- `KeepTogether As Byte  (read/write)`  
  Use the KeepTogether property for a group in a report to keep parts of a group (including the group header, detail section, and group footer) together on the same page. For example, you might want a group header to always be printed on the same page with the first detail section. Read/write Byte.
