# DateTime

**Type:** Module  
**Library:** Visual Basic For Applications  

## Functions (26)

- `_B_var_Date() As Variant`
- `_B_str_Date(Arg1 As String)`
- `_B_var_Date(Arg1 As Variant)`
- `_B_str_Date() As String`
- `DateSerial(Year As Integer, Month As Integer, Day As Integer) As Variant`  
  Returns a Variant (Date) for a specified year, month, and day.
    - `Year As Integer` (required): Required; Integer. Number between 100 and 9999, inclusive, or a numeric expression.
    - `Month As Integer` (required): Required; Integer. Any numeric expression.
    - `Day As Integer` (required): Required; Integer. Any numeric expression.
- `DateValue(Date As String) As Variant`  
  Returns a Variant (Date).
- `Day(Date As Variant) As Variant`  
  Returns a Variant (Integer) specifying a whole number between 1 and 31, inclusive, representing the day of the month.
- `Hour(Time As Variant) As Variant`  
  Returns a Variant (Integer) specifying a whole number between 0 and 23, inclusive, representing the hour of the day.
- `Minute(Time As Variant) As Variant`  
  Returns a Variant (Integer) specifying a whole number between 0 and 59, inclusive, representing the minute of the hour.
- `Month(Date As Variant) As Variant`  
  Returns a Variant (Integer) specifying a whole number between 1 and 12, inclusive, representing the month of the year.
- `Now() As Variant`  
  Returns a Variant (Date) specifying the current date and time according to your computer's system date and time.
- `Second(Time As Variant) As Variant`  
  Returns a Variant (Integer) specifying a whole number between 0 and 59, inclusive, representing the second of the minute.
- `_B_var_Time() As Variant`
- `_B_str_Time(Arg1 As String)`
- `_B_var_Time(Arg1 As Variant)`
- `_B_str_Time() As String`
- `Timer() As Single`  
  Returns a Single representing the number of seconds elapsed since midnight.
- `TimeSerial(Hour As Integer, Minute As Integer, Second As Integer) As Variant`  
  Returns a Variant (Date) containing the time for a specific hour, minute, and second.
    - `Hour As Integer` (required): Required; Variant (Integer). Number between 0 (12:00 A.M.) and 23 (11:00 P.M.), inclusive, or a numeric expression.
    - `Minute As Integer` (required): Required; Variant (Integer). Any numeric expression.
    - `Second As Integer` (required): Required; Variant (Integer). Any numeric expression.
- `TimeValue(Time As String) As Variant`  
  Returns a Variant (Date) containing the time.
- `Weekday(Date As Variant, [FirstDayOfWeek As VbDayOfWeek]) As Variant`  
  Returns a Variant (Integer) containing a whole number representing the day of the week.
    - `Date As Variant` (required): Required. Variant, numeric expression, string expression, or any combination, that can represent a date. If _date_ contains Null, Null is returned.
    - `FirstDayOfWeek As VbDayOfWeek` (optional): Optional. A constant that specifies the first day of the week. If not specified, vbSunday is assumed.
- `Year(Date As Variant) As Variant`  
  Returns a Variant (Integer) containing a whole number representing the year.
- `DateAdd(Interval As String, Number As Double, Date As Variant) As Variant`  
  Returns a Variant (Date) containing a date to which a specified time interval has been added.
    - `Interval As String` (required): Required. String expression that is the interval of time you want to add.
    - `Number As Double` (required): Required. Numeric expression that is the number of intervals you want to add. It can be positive (to get dates in the future) or negative (to get dates in the past).
    - `Date As Variant` (required): Required. Variant (Date) or literal representing the date to which the interval is added.
- `DateDiff(Interval As String, Date1 As Variant, Date2 As Variant, [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As Variant`  
  Returns a Variant (Long) specifying the number of time intervals between two specified dates.
    - `Interval As String` (required): Required. String expression that is the interval of time you use to calculate the difference between _date1_ and _date2_.
    - `FirstDayOfWeek As VbDayOfWeek` (optional): Optional. A constant that specifies the first day of the week. If not specified, Sunday is assumed.
    - `FirstWeekOfYear As VbFirstWeekOfYear` (optional): Optional. A constant that specifies the first week of the year. If not specified, the first week is assumed to be the week in which January 1 occurs.
- `DatePart(Interval As String, Date As Variant, [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As Variant`
- `Calendar() As VbCalendar`  
  Returns or sets a value specifying the type of calendar to use with your project.
- `Calendar(Arg1 As VbCalendar)`  
  Returns or sets a value specifying the type of calendar to use with your project.
