# Strings

**Type:** Module  
**Library:** Visual Basic For Applications  

## Functions (55)

- `Asc(String As String) As Integer`  
  Returns an Integer representing the character code corresponding to the first letter in a string.
- `_B_str_Chr(CharCode As Long) As String`
- `_B_var_Chr(CharCode As Long) As Variant`
- `_B_str_LCase(String As String) As String`
- `_B_var_LCase(String As Variant) As Variant`
- `_B_str_Mid(String As String, Start As Long, [Length As Variant]) As String`
- `_B_var_Mid(String As Variant, Start As Long, [Length As Variant]) As Variant`
- `_B_str_MidB(String As String, Start As Long, [Length As Variant]) As String`
- `_B_var_MidB(String As Variant, Start As Long, [Length As Variant]) As Variant`
- `InStr([Start As Variant], [String1 As Variant], [String2 As Variant], [Compare As VbCompareMethod]) As Variant`  
  Returns a Variant (Long) specifying the position of the first occurrence of one string within another.
    - `Start As Variant` (optional): Optional. Numeric expression that sets the starting position for each search. If omitted, search begins at the first character position. If _start_ contains Null, an error occurs. The _start_ argument is required if _compare_ is specified.
    - `String1 As Variant` (optional): Required. String expression being searched.
    - `String2 As Variant` (optional): Required. String expression sought.
    - `Compare As VbCompareMethod` (optional): Optional. Specifies the type of string comparison. If _compare_ is Null, an error occurs. If _compare_ is omitted, the Option Compare setting determines the type of comparison. Specify a valid LCID (LocaleID) to use locale-specific rules in the comparison.
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
  Returns a Variant (Integer) indicating the result of a string comparison.
    - `String1 As Variant` (required): Required. Any valid string expression.
    - `String2 As Variant` (required): Required. Any valid string expression.
    - `Compare As VbCompareMethod` (optional): Optional. Specifies the type of string comparison. If the _compare_ argument is Null, an error occurs. If _compare_ is omitted, the Option Compare setting determines the type of comparison.
- `_B_str_Format(Expression As Variant, [Format As Variant], [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As String`
- `_B_var_Format(Expression As Variant, [Format As Variant], [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As Variant`
- `Len(Expression As Variant) As Variant`  
  Returns a Long containing the number of characters in a string or the number of bytes required to store a variable.
- `LenB(Expression As Variant) As Variant`
- `AscB(String As String) As Byte`
- `_B_str_ChrB(CharCode As Byte) As String`
- `_B_var_ChrB(CharCode As Byte) As Variant`
- `AscW(String As String) As Integer`
- `_B_str_ChrW(CharCode As Long) As String`
- `_B_var_ChrW(CharCode As Long) As Variant`
- `FormatDateTime(Expression As Variant, [NamedFormat As VbDateTimeFormat]) As String`  
  Returns an expression formatted as a date or time.
    - `NamedFormat As VbDateTimeFormat` (optional): Optional. Numeric value that indicates the date/time format used. If omitted, vbGeneralDate is used.
- `FormatNumber(Expression As Variant, [NumDigitsAfterDecimal As Long], [IncludeLeadingDigit As VbTriState], [UseParensForNegativeNumbers As VbTriState], [GroupDigits As VbTriState]) As String`  
  Returns an expression formatted as a number.
    - `Expression As Variant` (required): Required. Expression to be formatted.
    - `NumDigitsAfterDecimal As Long` (optional): Optional. Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used.
    - `IncludeLeadingDigit As VbTriState` (optional): Optional. Tristate constant that indicates whether or not a leading zero is displayed for fractional values. See Settings section for values.
    - `UseParensForNegativeNumbers As VbTriState` (optional): Optional. Tristate constant that indicates whether or not to place negative values within parentheses. See Settings section for values.
    - `GroupDigits As VbTriState` (optional): Optional. Tristate constant that indicates whether or not numbers are grouped by using the group delimiter specified in the computer's regional settings. See Settings section for values.
- `FormatPercent(Expression As Variant, [NumDigitsAfterDecimal As Long], [IncludeLeadingDigit As VbTriState], [UseParensForNegativeNumbers As VbTriState], [GroupDigits As VbTriState]) As String`  
  Returns an expression formatted as a percentage (multipled by 100) with a trailing % character.
    - `Expression As Variant` (required): Required. Expression to be formatted.
    - `NumDigitsAfterDecimal As Long` (optional): Optional. Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used.
    - `IncludeLeadingDigit As VbTriState` (optional): Optional. Tristate constant that indicates whether or not a leading zero is displayed for fractional values. See Settings section for values.
    - `UseParensForNegativeNumbers As VbTriState` (optional): Optional. Tristate constant that indicates whether or not to place negative values within parentheses. See Settings section for values.
    - `GroupDigits As VbTriState` (optional): Optional. Tristate constant that indicates whether or not numbers are grouped by using the group delimiter specified in the computer's regional settings. See Settings section for values.
- `FormatCurrency(Expression As Variant, [NumDigitsAfterDecimal As Long], [IncludeLeadingDigit As VbTriState], [UseParensForNegativeNumbers As VbTriState], [GroupDigits As VbTriState]) As String`  
  Returns an expression formatted as a currency value by using the currency symbol defined in the system control panel.
    - `Expression As Variant` (required): Required. Expression to be formatted.
    - `NumDigitsAfterDecimal As Long` (optional): Optional. Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used.
    - `IncludeLeadingDigit As VbTriState` (optional): Optional. Tristate constant that indicates whether or not a leading zero is displayed for fractional values. See Settings section for values.
    - `UseParensForNegativeNumbers As VbTriState` (optional): Optional. Tristate constant that indicates whether or not to place negative values within parentheses. See Settings section for values.
    - `GroupDigits As VbTriState` (optional): Optional. Tristate constant that indicates whether or not numbers are grouped by using the group delimiter specified in the computer's regional settings. See Settings section for values.
- `WeekdayName(Weekday As Long, [Abbreviate As Boolean], [FirstDayOfWeek As VbDayOfWeek]) As String`  
  Returns a string indicating the specified day of the week.
    - `Weekday As Long` (required): Required. The numeric designation for the day of the week. Numeric value of each day depends on setting of the _firstdayofweek_ setting.
    - `Abbreviate As Boolean` (optional): Optional. Boolean value that indicates if the weekday name is to be abbreviated. If omitted, the default is False, which means that the weekday name is not abbreviated.
    - `FirstDayOfWeek As VbDayOfWeek` (optional): Optional. Numeric value indicating the first day of the week. See Settings section for values.
- `MonthName(Month As Long, [Abbreviate As Boolean]) As String`  
  Returns a string indicating the specified month.
    - `Month As Long` (required): Required. The numeric designation of the month. For example, January is 1, February is 2, and so on.
    - `Abbreviate As Boolean` (optional): Optional. Boolean value that indicates if the month name is to be abbreviated. If omitted, the default is False, which means that the month name is not abbreviated.
- `Replace(Expression As String, Find As String, Replace As String, [Start As Long], [Count As Long], [Compare As VbCompareMethod]) As String`  
  Returns a string, which is a substring of a string expression beginning at the start position (defaults to 1), in which a specified substring has been replaced with another substring a specified number of times.
    - `Expression As String` (required): Required. String expression containing substring to replace.
    - `Find As String` (required): Required. Substring being searched for.
    - `Replace As String` (required): Required. Replacement substring.
    - `Start As Long` (optional): Optional. Start position for the substring of _expression_ to be searched and returned. If omitted, 1 is assumed.
    - `Count As Long` (optional): Optional. Number of substring substitutions to perform. If omitted, the default value is -1, which means, make all possible substitutions.
    - `Compare As VbCompareMethod` (optional): Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. See Settings section for values.
- `StrReverse(Expression As String) As String`  
  Returns a string in which the character order of a specified string is reversed.
- `Join(SourceArray As Variant, [Delimiter As Variant]) As String`  
  Returns a string created by joining a number of substrings contained in an array.
    - `SourceArray As Variant` (required): Required. One-dimensional array containing substrings to be joined.
    - `Delimiter As Variant` (optional): Optional. String character used to separate the substrings in the returned string. If omitted, the space character (" ") is used. If _delimiter_ is a zero-length string (""), all items in the list are concatenated with no delimiters.
- `Filter(SourceArray As Variant, Match As String, [Include As Boolean], [Compare As VbCompareMethod]) As Variant`  
  Returns a zero-based array containing a subset of a string array based on a specified filter criteria.
    - `SourceArray As Variant` (required): Required. One-dimensional array of strings to be searched.
    - `Match As String` (required): Required. String to search for.
    - `Include As Boolean` (optional): Optional. Boolean value indicating whether to return substrings that include or exclude _match_. If _include_ is True, Filter returns the subset of the array that contains _match_ as a substring. If _include_ is False, Filter returns the subset of the array that does not contain _match_ as a substring.
    - `Compare As VbCompareMethod` (optional): Optional. Numeric value indicating the kind of string comparison to use. See Settings section for values.
- `InStrRev(StringCheck As String, StringMatch As String, [Start As Long], [Compare As VbCompareMethod]) As Long`  
  Returns the position of an occurrence of one string within another, from the end of the string.
    - `StringCheck As String` (required): Required. String expression being searched.
    - `StringMatch As String` (required): Required. String expression being searched for.
    - `Start As Long` (optional): Optional. Numeric expression that sets the starting position for each search. If omitted, -1 is used, which means that the search begins at the last character position. If _start_ contains Null, an error occurs.
    - `Compare As VbCompareMethod` (optional): Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. If omitted, a binary comparison is performed. See the Settings section for values.
- `Split(Expression As String, [Delimiter As Variant], [Limit As Long], [Compare As VbCompareMethod]) As Variant`  
  Returns a zero-based, one-dimensional array containing a specified number of substrings.
    - `Expression As String` (required): Required. String expression containing substrings and delimiters. If _expression_ is a zero-length string(""), Split returns an empty array, that is, an array with no elements and no data.
    - `Delimiter As Variant` (optional): Optional. String character used to identify substring limits. If omitted, the space character (" ") is assumed to be the delimiter. If _delimiter_ is a zero-length string, a single-element array containing the entire _expression_ string is returned.
    - `Limit As Long` (optional): Optional. Number of substrings to be returned; -1 indicates that all substrings are returned.
    - `Compare As VbCompareMethod` (optional): Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. See Settings section for values.
