# WdFieldType

**Type:** Enumeration  
**Library:** Microsoft Word 16.0 Object Library  

Specifies a Microsoft Word field. Unless otherwise specified, the field types described in this enumeration can be added interactively to a Word document by using the Field dialog box. See the Word Help for more information about specific field codes.

## Constants (97)

- `wdFieldEmpty` = -1  
  Empty field. Acts as a placeholder for field content that has not yet been added. A field added by pressing Ctrl+F9 in the user interface is an Empty field.
- `wdFieldRef` = 3  
  Ref field.
- `wdFieldIndexEntry` = 4  
  XE (Index Entry) field.
- `wdFieldFootnoteRef` = 5  
  FootnoteRef field. Not available through the Field dialog box. Inserted programmatically or interactively.
- `wdFieldSet` = 6  
  Set field.
- `wdFieldIf` = 7  
  If field.
- `wdFieldIndex` = 8  
  Index field.
- `wdFieldTOCEntry` = 9  
  TOC (Table of Contents Entry) field.
- `wdFieldStyleRef` = 10  
  StyleRef field.
- `wdFieldRefDoc` = 11  
  RD (Reference Document) field.
- `wdFieldSequence` = 12  
  Seq (Sequence) field.
- `wdFieldTOC` = 13  
  TOC (Table of Contents) field.
- `wdFieldInfo` = 14  
  Info field.
- `wdFieldTitle` = 15  
  Title field.
- `wdFieldSubject` = 16  
  Subject field.
- `wdFieldAuthor` = 17  
  Author field.
- `wdFieldKeyWord` = 18  
  Keywords field.
- `wdFieldComments` = 19  
  Comments field.
- `wdFieldLastSavedBy` = 20  
  LastSavedBy field.
- `wdFieldCreateDate` = 21  
  CreateDate field.
- `wdFieldSaveDate` = 22  
  SaveDate field.
- `wdFieldPrintDate` = 23  
  PrintDate field.
- `wdFieldRevisionNum` = 24  
  RevNum field.
- `wdFieldEditTime` = 25  
  EditTime field.
- `wdFieldNumPages` = 26  
  NumPages field.
- `wdFieldNumWords` = 27  
  NumWords field.
- `wdFieldNumChars` = 28  
  NumChars field.
- `wdFieldFileName` = 29  
  FileName field.
- `wdFieldTemplate` = 30  
  Template field.
- `wdFieldDate` = 31  
  Date field.
- `wdFieldTime` = 32  
  Time field.
- `wdFieldPage` = 33  
  Page field.
- `wdFieldExpression` = 34  
  = (Formula) field.
- `wdFieldQuote` = 35  
  Quote field.
- `wdFieldInclude` = 36  
  Include field. Cannot be added through the Field dialog box, but can be added interactively or through code.
- `wdFieldPageRef` = 37  
  PageRef field.
- `wdFieldAsk` = 38  
  Ask field.
- `wdFieldFillIn` = 39  
  Fill-In field.
- `wdFieldData` = 40  
  Data field.
- `wdFieldNext` = 41  
  Next field.
- `wdFieldNextIf` = 42  
  NextIf field.
- `wdFieldSkipIf` = 43  
  SkipIf field.
- `wdFieldMergeRec` = 44  
  MergeRec field.
- `wdFieldDDE` = 45  
  DDE field. No longer available through the Field dialog box, but supported for documents created in earlier versions of Word.
- `wdFieldDDEAuto` = 46  
  DDEAuto field. No longer available through the Field dialog box, but supported for documents created in earlier versions of Word.
- `wdFieldGlossary` = 47  
  Glossary field. No longer supported in Word.
- `wdFieldPrint` = 48  
  Print field.
- `wdFieldFormula` = 49  
  EQ (Equation) field.
- `wdFieldGoToButton` = 50  
  GoToButton field.
- `wdFieldMacroButton` = 51  
  MacroButton field.
- `wdFieldAutoNumOutline` = 52  
  AutoNumOut field.
- `wdFieldAutoNumLegal` = 53  
  AutoNumLgl field.
- `wdFieldAutoNum` = 54  
  AutoNum field.
- `wdFieldImport` = 55  
  Import field. Cannot be added through the Field dialog box, but can be added interactively or through code.
- `wdFieldLink` = 56  
  Link field.
- `wdFieldSymbol` = 57  
  Symbol field.
- `wdFieldEmbed` = 58  
  Embedded field.
- `wdFieldMergeField` = 59  
  MergeField field.
- `wdFieldUserName` = 60  
  UserName field.
- `wdFieldUserInitials` = 61  
  UserInitials field.
- `wdFieldUserAddress` = 62  
  UserAddress field.
- `wdFieldBarCode` = 63  
  BarCode field.
- `wdFieldDocVariable` = 64  
  DocVariable field.
- `wdFieldSection` = 65  
  Section field.
- `wdFieldSectionPages` = 66  
  SectionPages field.
- `wdFieldIncludePicture` = 67  
  IncludePicture field.
- `wdFieldIncludeText` = 68  
  IncludeText field.
- `wdFieldFileSize` = 69  
  FileSize field.
- `wdFieldFormTextInput` = 70  
  FormText field.
- `wdFieldFormCheckBox` = 71  
  FormCheckBox field.
- `wdFieldNoteRef` = 72  
  NoteRef field.
- `wdFieldTOA` = 73  
  TOA (Table of Authorities) field.
- `wdFieldTOAEntry` = 74  
  TOA (Table of Authorities Entry) field.
- `wdFieldMergeSeq` = 75  
  MergeSeq field.
- `wdFieldPrivate` = 77  
  Private field.
- `wdFieldDatabase` = 78  
  Database field.
- `wdFieldAutoText` = 79  
  AutoText field.
- `wdFieldCompare` = 80  
  Compare field.
- `wdFieldAddin` = 81  
  Add-in field. Not available through the Field dialog box. Used to store data that is hidden from the user interface.
- `wdFieldSubscriber` = 82  
  Macintosh only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `wdFieldFormDropDown` = 83  
  FormDropDown field.
- `wdFieldAdvance` = 84  
  Advance field.
- `wdFieldDocProperty` = 85  
  DocProperty field.
- `wdFieldOCX` = 87  
  OCX field. Cannot be added through the Field dialog box, but can be added through code by using the AddOLEControl method of the Shapes collection or of the InlineShapes collection.
- `wdFieldHyperlink` = 88  
  Hyperlink field.
- `wdFieldAutoTextList` = 89  
  AutoTextList field.
- `wdFieldListNum` = 90  
  ListNum field.
- `wdFieldHTMLActiveX` = 91  
  HTMLActiveX field. Not currently supported.
- `wdFieldBidiOutline` = 92  
  BidiOutline field.
- `wdFieldAddressBlock` = 93  
  AddressBlock field.
- `wdFieldGreetingLine` = 94  
  GreetingLine field.
- `wdFieldShape` = 95  
  Shape field. Automatically created for any drawn picture.
- `wdFieldCitation` = 96  
  Citation field.
- `wdFieldBibliography` = 97  
  Bibliography field.
- `wdFieldMergeBarcode` = 98  
  MergeBarcode field.
- `wdFieldDisplayBarcode` = 99  
  DisplayBarcode field.
- `wdFieldCopilot` = 100
