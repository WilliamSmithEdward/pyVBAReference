# Strings

**Type:** Module  
**Library:** Visual Basic For Applications  

## Functions (55)

- `Asc(String As String) As Integer`
- `_B_str_Chr(CharCode As Long) As String`
- `_B_var_Chr(CharCode As Long) As Variant`
- `_B_str_LCase(String As String) As String`
- `_B_var_LCase(String As Variant) As Variant`
- `_B_str_Mid(String As String, Start As Long, [Length As Variant]) As String`
- `_B_var_Mid(String As Variant, Start As Long, [Length As Variant]) As Variant`
- `_B_str_MidB(String As String, Start As Long, [Length As Variant]) As String`
- `_B_var_MidB(String As Variant, Start As Long, [Length As Variant]) As Variant`
- `InStr([Start As Variant], [String1 As Variant], [String2 As Variant], [Compare As VbCompareMethod]) As Variant`
- `InStrB([Start As Variant], [String1 As Variant], [String2 As Variant], [Compare As VbCompareMethod]) As Variant`
- `_B_str_Left(String As String, Length As Long) As String`
- `_B_var_Left(String As Variant, Length As Long) As Variant`
- `_B_str_LeftB(String As String, Length As Long) As String`
- `_B_var_LeftB(String As Variant, Length As Long) As Variant`
- `_B_str_LTrim(String As String) As String`
- `_B_var_LTrim(String As Variant) As Variant`
- `_B_str_RightB(String As String, Length As Long) As String`
- `_B_var_RightB(String As Variant, Length As Long) As Variant`
- `_B_str_Right(String As String, Length As Long) As String`
- `_B_var_Right(String As Variant, Length As Long) As Variant`
- `_B_str_RTrim(String As String) As String`
- `_B_var_RTrim(String As Variant) As Variant`
- `_B_str_Space(Number As Long) As String`
- `_B_var_Space(Number As Long) As Variant`
- `_B_var_StrConv(String As Variant, Conversion As VbStrConv, [LocaleID As Long]) As Variant`
- `_B_str_String(Number As Long, Character As Variant) As String`
- `_B_var_String(Number As Long, Character As Variant) As Variant`
- `_B_str_Trim(String As String) As String`
- `_B_var_Trim(String As Variant) As Variant`
- `_B_str_UCase(String As String) As String`
- `_B_var_UCase(String As Variant) As Variant`
- `StrComp(String1 As Variant, String2 As Variant, [Compare As VbCompareMethod]) As Variant`
- `_B_str_Format(Expression As Variant, [Format As Variant], [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As String`
- `_B_var_Format(Expression As Variant, [Format As Variant], [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As Variant`
- `Len(Expression As Variant) As Variant`
- `LenB(Expression As Variant) As Variant`
- `AscB(String As String) As Byte`
- `_B_str_ChrB(CharCode As Byte) As String`
- `_B_var_ChrB(CharCode As Byte) As Variant`
- `AscW(String As String) As Integer`
- `_B_str_ChrW(CharCode As Long) As String`
- `_B_var_ChrW(CharCode As Long) As Variant`
- `FormatDateTime(Expression As Variant, [NamedFormat As VbDateTimeFormat]) As String`  
  Formats expression as date
- `FormatNumber(Expression As Variant, [NumDigitsAfterDecimal As Long], [IncludeLeadingDigit As VbTriState], [UseParensForNegativeNumbers As VbTriState], [GroupDigits As VbTriState]) As String`  
  Formats expression as number
- `FormatPercent(Expression As Variant, [NumDigitsAfterDecimal As Long], [IncludeLeadingDigit As VbTriState], [UseParensForNegativeNumbers As VbTriState], [GroupDigits As VbTriState]) As String`  
  Formats expression as percent
- `FormatCurrency(Expression As Variant, [NumDigitsAfterDecimal As Long], [IncludeLeadingDigit As VbTriState], [UseParensForNegativeNumbers As VbTriState], [GroupDigits As VbTriState]) As String`  
  Formats expression as currency
- `WeekdayName(Weekday As Long, [Abbreviate As Boolean], [FirstDayOfWeek As VbDayOfWeek]) As String`  
  Returns localized weekday name
- `MonthName(Month As Long, [Abbreviate As Boolean]) As String`  
  Returns localized month name
- `Replace(Expression As String, Find As String, Replace As String, [Start As Long], [Count As Long], [Compare As VbCompareMethod]) As String`  
  Find and replace a substring within a string
- `StrReverse(Expression As String) As String`  
  Reverse a string
- `Join(SourceArray As Variant, [Delimiter As Variant]) As String`  
  Join array elements into a string
- `Filter(SourceArray As Variant, Match As String, [Include As Boolean], [Compare As VbCompareMethod]) As Variant`  
  Return array of matches
- `InStrRev(StringCheck As String, StringMatch As String, [Start As Long], [Compare As VbCompareMethod]) As Long`  
  Returns the position of the last occurrence of one string within another
- `Split(Expression As String, [Delimiter As Variant], [Limit As Long], [Compare As VbCompareMethod]) As Variant`  
  Split a string into an array
