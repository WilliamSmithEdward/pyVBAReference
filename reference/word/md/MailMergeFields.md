# MailMergeFields

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002091F-0000-0000-C000-000000000046}  

A collection of MailMergeField objects that represent the mail merge related fields in a document.

**Remarks:** Use the Fields property to return the MailMergeFields collection. The following example adds an ASK field after the last mail merge field in the active document. Use the Add method to add a merge field to the MailMergeFields collection. The following example replaces the selection with a MiddleInitial merge field. Use Fields (Index), where Index is the index number, to return a single MailMergeField object. The following example displays the field code of the first mail merge field in the active document. The MailMergeFields collection has additional methods, such as AddAsk and AddFillIn, for adding fields related to a mail merge operation.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MailMergeFields object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of mail merge fields in the collection. Read-only.

## Methods (11)

- `Item(Index As Long) As MailMergeField`  
  Returns an individual MailMergeField object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add(Range As Range, Name As String) As MailMergeField`  
  Returns a MailMergeField object that represents a mail merge field added to the data source document.
    - `Range As Range` (required): The range where you want the field to appear. This field replaces the range, if the range isn't collapsed.
    - `Name As String` (required): The name of the field.
- `AddAsk(Range As Range, Name As String, [Prompt As Variant], [DefaultAskText As Variant], [AskOnce As Variant]) As MailMergeField`  
  Adds an ASK field to a mail merge main document. Returns a MailMergeField object.
    - `Range As Range` (required): The location for the ASK field.
    - `Name As String` (required): The bookmark name that the response or default text is assigned to. Use a REF field with the bookmark name to display the result in a document.
    - `Prompt As Variant` (optional): The text that's displayed in the dialog box.
    - `DefaultAskText As Variant` (optional): The default response, which appears in the text box when the dialog box is displayed. Corresponds to the \d switch for an ASK field.
    - `AskOnce As Variant` (optional): True to display the dialog box only once instead of each time a new record is merged. Corresponds to the \o switch for an ASK field.
- `AddFillIn(Range As Range, [Prompt As Variant], [DefaultFillInText As Variant], [AskOnce As Variant]) As MailMergeField`  
  Adds a FILLIN field to a mail merge main document. Returns a MailMergeField object.
    - `Range As Range` (required): The location for the FILLIN field.
    - `Prompt As Variant` (optional): The text that's displayed in the dialog box.
    - `DefaultFillInText As Variant` (optional): The default response, which appears in the text box when the dialog box is displayed. Corresponds to the \d switch for an FILLIN field.
    - `AskOnce As Variant` (optional): True to display the prompt only once instead of each time a new record is merged. Corresponds to the \o switch for a FILLIN field. The default value is False.
- `AddIf(Range As Range, MergeField As String, Comparison As WdMailMergeComparison, [CompareTo As Variant], [TrueAutoText As Variant], [TrueText As Variant], [FalseAutoText As Variant], [FalseText As Variant]) As MailMergeField`  
  Adds an IF field to a mail merge main document. Returns a MailMergeField object.
    - `Range As Range` (required): The location for the IF field.
    - `MergeField As String` (required): The merge field name.
    - `Comparison As WdMailMergeComparison` (required): The operator used in the comparison.
    - `CompareTo As Variant` (optional): The text to compare with the contents of MergeField.
    - `TrueAutoText As Variant` (optional): The AutoText entry that's inserted if the comparison is true. If this argument is specified, TrueText is ignored.
    - `TrueText As Variant` (optional): The text that's inserted if the comparison is true.
    - `FalseAutoText As Variant` (optional): The AutoText entry that's inserted if the comparison is false. If this argument is specified, FalseText is ignored.
    - `FalseText As Variant` (optional): The text that's inserted if the comparison is false.
- `AddMergeRec(Range As Range) As MailMergeField`  
  Adds a MERGEREC field to a mail merge main document. Returns a MailMergeField object.
    - `Range As Range` (required): The location for the MERGEREC field.
- `AddMergeSeq(Range As Range) As MailMergeField`  
  Adds a MERGESEQ field to a mail merge main document. Returns a MailMergeField object.
    - `Range As Range` (required): The location for the MERGESEQ field.
- `AddNext(Range As Range) As MailMergeField`  
  Adds a NEXT field to a mail merge main document. Returns a MailMergeField object.
    - `Range As Range` (required): The location for the NEXT field.
- `AddNextIf(Range As Range, MergeField As String, Comparison As WdMailMergeComparison, [CompareTo As Variant]) As MailMergeField`  
  Adds a NEXTIF field to a mail merge main document. Returns a MailMergeField object.
    - `Range As Range` (required): The location for the NEXTIF field.
    - `MergeField As String` (required): The merge field name.
    - `Comparison As WdMailMergeComparison` (required): The operator used in the comparison.
    - `CompareTo As Variant` (optional): The text to compare with the contents of MergeField.
- `AddSet(Range As Range, Name As String, [ValueText As Variant], [ValueAutoText As Variant]) As MailMergeField`  
  Adds a SET field to a mail merge main document. Returns a MailMergeField object.
    - `Range As Range` (required): The location for the SET field.
    - `Name As String` (required): The bookmark name that ValueText is assigned to.
    - `ValueText As Variant` (optional): The text associated with the bookmark specified by the Name argument.
    - `ValueAutoText As Variant` (optional): The AutoText entry that includes text associated with the bookmark specified by the Name argument. If this argument is specified, ValueText is ignored.
- `AddSkipIf(Range As Range, MergeField As String, Comparison As WdMailMergeComparison, [CompareTo As Variant]) As MailMergeField`  
  Adds a SKIPIF field to a mail merge main document. Returns a MailMergeField object. .
    - `Range As Range` (required): The location for the SKIPIF field.
    - `MergeField As String` (required): The merge field name.
    - `Comparison As WdMailMergeComparison` (required): The operator used in the comparison.
    - `CompareTo As Variant` (optional): The text to compare with the contents of MergeField.
