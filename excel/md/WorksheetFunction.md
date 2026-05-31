# WorksheetFunction

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020845-0000-0000-C000-000000000046}  

Used as a container for Microsoft Excel worksheet functions that can be called from Visual Basic.

**Example:**

```vba
Set myRange = Worksheets("Sheet1").Range("A1:C10")
answer = Application.WorksheetFunction.Min(myRange)
MsgBox answer
```

## Properties (3)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.

## Methods (403)

- `Count(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Counts the number of cells that contain numbers and counts numbers within the list of arguments.
- `IsNA(Arg1 As Variant) As Boolean`  
  Checks the type of value and returns True or False depending on whether the value refers to the #N/A (value not available) error value.
    - `Arg1 As Variant` (required): Value - the value that you want tested. Value can be a blank (empty cell), error, logical, text, number, or reference value, or a name referring to any of these, that you want to test.
- `IsError(Arg1 As Variant) As Boolean`  
  Checks the type of value and returns True or False depending on whether the value refers to any error value (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!).
    - `Arg1 As Variant` (required): Value - the value that you want tested. Value can be a blank (empty cell), error, logical, text, number, or reference value, or a name referring to any of these, that you want to test.
- `Sum(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Adds all the numbers in a range of cells.
- `Average(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the average (arithmetic mean) of the arguments.
- `Min(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the smallest number in a set of values.
- `Max(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the largest value in a set of values.
- `Npv(Arg1 As Double, Arg2 As Variant, [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Calculates the net present value of an investment by using a discount rate and a series of future payments (negative values) and income (positive values).
    - `Arg1 As Double` (required): Rate - the rate of discount over the length of one period.
- `StDev(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Estimates standard deviation based on a sample. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).
- `Dollar(Arg1 As Double, [Arg2 As Variant]) As String`  
  The function described in this Help topic converts a number to text format and applies a currency symbol. The name of the function (and the symbol that it applies) depends upon your language settings.
    - `Arg1 As Double` (required): Number - a number, a reference to a cell containing a number, or a formula that evaluates to a number.
    - `Arg2 As Variant` (optional): Decimals - the number of digits to the right of the decimal point. If decimals is negative, number is rounded to the left of the decimal point. If you omit decimals, it is assumed to be 2.
- `Fixed(Arg1 As Double, [Arg2 As Variant], [Arg3 As Variant]) As String`  
  Rounds a number to the specified number of decimals, formats the number in decimal format using a period and commas, and returns the result as text.
    - `Arg1 As Double` (required): Number - the number you want to round and convert to text.
    - `Arg2 As Variant` (optional): Decimals - the number of digits to the right of the decimal point.
    - `Arg3 As Variant` (optional): No_commas - a logical value that, if True, prevents Fixed from including commas in the returned text.
- `Pi() As Double`  
  Returns the number 3.14159265358979, the mathematical constant pi, accurate to 15 digits.
- `Ln(Arg1 As Double) As Double`  
  Returns the natural logarithm of a number. Natural logarithms are based on the constant e (2.71828182845904).
    - `Arg1 As Double` (required): Number - the positive real number for which you want the natural logarithm.
- `Log10(Arg1 As Double) As Double`  
  Returns the base-10 logarithm of a number.
    - `Arg1 As Double` (required): Number - the positive real number for which you want the base-10 logarithm.
- `Round(Arg1 As Double, Arg2 As Double) As Double`  
  Rounds a number to a specified number of digits.
    - `Arg1 As Double` (required): Number - the number that you want to round.
    - `Arg2 As Double` (required): Num_digits - specifies the number of digits to which you want to round number.
- `Lookup(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As Variant`  
  Returns a value either from a one-row or one-column range or from an array. The Lookup function has two syntax forms: the vector form and the array form.
    - `Arg1 As Variant` (required): Lookup_value - A value that Lookup searches for in the first vector. Lookup_value can be a number, text, a logical value, or a name or reference that refers to a value.
    - `Arg2 As Variant` (required): Lookup_vector or Array - In vector form, a range that contains only one row or one column. The values in lookup_vector can be text, numbers, or logical values. In array form, a range of cells that contains text, numbers, or logical values that you want to compare with lookup_value.
    - `Arg3 As Variant` (optional): Result_vector - Only used with the vector form. A range that contains only one row or column. It must be the same size as lookup_vector.
- `Index(Arg1 As Variant, Arg2 As Double, [Arg3 As Variant], [Arg4 As Variant]) As Variant`  
  Returns a value or the reference to a value from within a table or range. There are two forms of the Index function: the array form and the reference form.
    - `Arg1 As Variant` (required): Array or Reference - a range of cells or an array constant. For references, it is the reference to one or more cell ranges.
    - `Arg2 As Double` (required): Row_num - selects the row in array from which to return a value. If row_num is omitted, column_num is required. For references, the number of the row in reference from which to return a reference.
    - `Arg3 As Variant` (optional): Column_num - selects the column in array from which to return a value. If column_num is omitted, row_num is required. For reference, the number of the column in reference from which to return a reference.
    - `Arg4 As Variant` (optional): Area_num - only used when returning references. Selects a range in reference from which to return the intersection of row_num and column_num. The first area selected or entered is numbered 1, the second is 2, and so on. If area_num is omitted, Index uses area 1.
- `Rept(Arg1 As String, Arg2 As Double) As String`  
  Repeats text a given number of times. Use Rept to fill a cell with a number of instances of a text string.
    - `Arg1 As String` (required): Text - the text that you want to repeat.
    - `Arg2 As Double` (required): Number_times - a positive number specifying the number of times to repeat text.
- `And(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Boolean`  
  Returns True if all its arguments are True; returns False if one or more arguments is False.
- `Or(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Boolean`  
  Returns True if any argument is True; returns False if all arguments are False.
    - `Arg1 As Variant` (required): Logical1, logical2, ... - 1 to 30 conditions that you want to test that can be either True or False.
- `DCount(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Double`  
  Counts the cells that contain numbers in a column of a list or database that match conditions that you specify.
    - `Arg1 As Range` (required): Database - the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Field - indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): Criteria - the range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as the argument includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `DSum(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Double`  
  Adds the numbers in a column of a list or database that match conditions that you specify.
    - `Arg1 As Range` (required): Database - the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Field - indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): Criteria - the range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `DAverage(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Double`  
  Averages the values in a column of a list or database that match conditions that you specify.
    - `Arg1 As Range` (required): The range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): The range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `DMin(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Double`  
  Returns the smallest number in a column of a list or database that matches conditions that you specify.
    - `Arg1 As Range` (required): Database - the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Field - indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): Criteria - the range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `DMax(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Double`  
  Returns the largest number in a column of a list or database that matches conditions that you specify.
    - `Arg1 As Range` (required): Database - the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Field - indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): Criteria - the range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `DStDev(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Double`  
  Estimates the standard deviation of a population based on a sample by using the numbers in a column of a list or database that match conditions that you specify.
    - `Arg1 As Range` (required): Database - the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Field - indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): Criteria - the range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `Var(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Estimates variance based on a sample.
- `DVar(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Double`  
  Estimates the variance of a population based on a sample by using the numbers in a column of a list or database that match conditions that you specify.
    - `Arg1 As Range` (required): Database - the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Field - indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): Criteria - the range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `Text(Arg1 As Variant, Arg2 As String) As String`  
  Converts a value to text in a specific number format.
    - `Arg1 As Variant` (required): A numeric value, a formula that evaluates to a numeric value, or a reference to a cell containing a numeric value.
    - `Arg2 As String` (required): A number format in text form in the Category box on the Number tab in the Format Cells dialog box.
- `LinEst(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant]) As Variant`  
  Calculates the statistics for a line by using the least squares method to calculate a straight line that best fits your data, and returns an array that describes the line. Because this function returns an array of values, it must be entered as an array formula.
    - `Arg1 As Variant` (required): Known_y's - the set of y-values that you already know in the relationship y = mx + b.
    - `Arg2 As Variant` (optional): Known_x's - an optional set of x-values that you may already know in the relationship y = mx + b.
    - `Arg3 As Variant` (optional): Const - a logical value specifying whether to force the constant b to equal 0.
    - `Arg4 As Variant` (optional): Stats - a logical value specifying whether to return additional regression statistics.
- `Trend(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant]) As Variant`  
  Returns values along a linear trend. Fits a straight line (using the method of least squares) to the arrays known_y's and known_x's. Returns the y-values along that line for the array of new_x's that you specify.
    - `Arg1 As Variant` (required): Known_y's - the set of y-values that you already know in the relationship y = mx + b.
    - `Arg2 As Variant` (optional): Known_x's - an optional set of x-values that you may already know in the relationship y = mx + b.
    - `Arg3 As Variant` (optional): New_x's - new x-values for which you want Trend to return corresponding y-values.
    - `Arg4 As Variant` (optional): Const - a logical value specifying whether to force the constant b to equal 0.
- `LogEst(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant]) As Variant`  
  In regression analysis, calculates an exponential curve that fits your data, and returns an array of values that describes the curve. Because this function returns an array of values, it must be entered as an array formula.
    - `Arg1 As Variant` (required): Known_y's - the set of y-values that you already know in the relationship y = b*m^x.
    - `Arg2 As Variant` (optional): Known_x's - an optional set of x-values that you may already know in the relationship y = b*m^x.
    - `Arg3 As Variant` (optional): Const - a logical value specifying whether to force the constant b to equal 1.
    - `Arg4 As Variant` (optional): Stats - a logical value specifying whether to return additional regression statistics.
- `Growth(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant]) As Variant`  
  Calculates predicted exponential growth by using existing data. Growth returns the y-values for a series of new x-values that you specify by using existing x-values and y-values. You can also use the Growth worksheet function to fit an exponential curve to existing x-values and y-values.
    - `Arg1 As Variant` (required): Known_y's - the set of y-values that you already know in the relationship y = b*m^x.
    - `Arg2 As Variant` (optional): Known_x's - an optional set of x-values that you may already know in the relationship y = b*m^x.
    - `Arg3 As Variant` (optional): New_x's - new x-values for which you want Growth to return corresponding y-values.
    - `Arg4 As Variant` (optional): Const - a logical value specifying whether to force the constant b to equal 1.
- `Pv(Arg1 As Double, Arg2 As Double, Arg3 As Double, [Arg4 As Variant], [Arg5 As Variant]) As Double`  
  Returns the present value of an investment. The present value is the total amount that a series of future payments is worth now. For example, when you borrow money, the loan amount is the present value to the lender.
    - `Arg1 As Double` (required): Rate - the interest rate per period. For example, if you obtain an automobile loan at a 10 percent annual interest rate and make monthly payments, your interest rate per month is 10%/12, or 0.83%. You would enter 10%/12, or 0.83%, or 0.0083, into the formula as the rate.
    - `Arg2 As Double` (required): Nper - the total number of payment periods in an annuity. For example, if you get a four-year car loan and make monthly payments, your loan has 4*12 (or 48) periods. You would enter 48 into the formula for nper.
    - `Arg3 As Double` (required): Pmt - the payment made each period and cannot change over the life of the annuity. Typically, pmt includes principal and interest but no other fees or taxes. For example, the monthly payments on a $10,000, four-year car loan at 12 percent are $263.33. You would enter -263.33 into the formula as the pmt. If pmt is omitted, you must include the fv argument.
    - `Arg4 As Variant` (optional): Fv - the future value, or a cash balance you want to attain after the last payment is made. If fv is omitted, it is assumed to be 0 (the future value of a loan, for example, is 0). For example, if you want to save $50,000 to pay for a special project in 18 years, $50,000 is the future value. You could then make a conservative guess at an interest rate and determine how much you must save each month. If fv is omitted, you must include the pmt argument.
    - `Arg5 As Variant` (optional): Type - the number 0 or 1; indicates when payments are due.
- `Fv(Arg1 As Double, Arg2 As Double, Arg3 As Double, [Arg4 As Variant], [Arg5 As Variant]) As Double`  
  Returns the future value of an investment based on periodic, constant payments and a constant interest rate.
    - `Arg1 As Double` (required): Rate - the interest rate per period.
    - `Arg2 As Double` (required): Nper - the total number of payment periods in an annuity.
    - `Arg3 As Double` (required): Pmt - the payment made each period; it cannot change over the life of the annuity. Typically, pmt contains principal and interest but no other fees or taxes. If pmt is omitted, you must include the pv argument.
    - `Arg4 As Variant` (optional): Pv - the present value, or the lump-sum amount that a series of future payments is worth right now. If pv is omitted, it is assumed to be 0 (zero), and you must include the pmt argument.
    - `Arg5 As Variant` (optional): Type - the number 0 or 1 and indicates when payments are due. If type is omitted, it is assumed to be 0.
- `NPer(Arg1 As Double, Arg2 As Double, Arg3 As Double, [Arg4 As Variant], [Arg5 As Variant]) As Double`  
  Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.
    - `Arg1 As Double` (required): Rate - the interest rate per period.
    - `Arg2 As Double` (required): Pmt - the payment made each period; it cannot change over the life of the annuity. Typically, pmt contains principal and interest but no other fees or taxes.
    - `Arg3 As Double` (required): Pv - the present value, or the lump-sum amount that a series of future payments is worth right now.
    - `Arg4 As Variant` (optional): Fv - the future value, or a cash balance that you want to attain after the last payment is made. If fv is omitted, it is assumed to be 0 (the future value of a loan, for example, is 0).
    - `Arg5 As Variant` (optional): Type - the number 0 or 1 and indicates when payments are due.
- `Pmt(Arg1 As Double, Arg2 As Double, Arg3 As Double, [Arg4 As Variant], [Arg5 As Variant]) As Double`  
  Calculates the payment for a loan based on constant payments and a constant interest rate.
    - `Arg1 As Double` (required): Rate - the interest rate for the loan.
    - `Arg2 As Double` (required): Nper - the total number of payments for the loan.
    - `Arg3 As Double` (required): Pv - the present value, or the total amount that a series of future payments is worth now; also known as the principal.
    - `Arg4 As Variant` (optional): Fv - the future value, or a cash balance you want to attain after the last payment is made. If fv is omitted, it is assumed to be 0 (zero), that is, the future value of a loan is 0.
    - `Arg5 As Variant` (optional): Type - the number 0 (zero) or 1; indicates when payments are due.
- `Rate(Arg1 As Double, Arg2 As Double, Arg3 As Double, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant]) As Double`  
  Returns the interest rate per period of an annuity. Rate is calculated by iteration and can have zero or more solutions. If the successive results of Rate don't converge to within 0.0000001 after 20 iterations, Rate returns the #NUM! error value.
    - `Arg1 As Double` (required): Nper - the total number of payment periods in an annuity.
    - `Arg2 As Double` (required): Pmt - the payment made each period and cannot change over the life of the annuity. Typically, pmt includes principal and interest but no other fees or taxes. If pmt is omitted, you must include the fv argument.
    - `Arg3 As Double` (required): Pv - the present value&mdash;the total amount that a series of future payments is worth now.
    - `Arg4 As Variant` (optional): Fv - the future value, or a cash balance you want to attain after the last payment is made. If fv is omitted, it is assumed to be 0 (the future value of a loan, for example, is 0).
    - `Arg5 As Variant` (optional): Type - the number 0 or 1; indicates when payments are due.
    - `Arg6 As Variant` (optional): Guess - your guess for what the rate will be.
- `MIrr(Arg1 As Variant, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the modified internal rate of return for a series of periodic cash flows. MIrr considers both the cost of the investment and the interest received on reinvestment of cash.
    - `Arg1 As Variant` (required): Values - an array or a reference to cells that contain numbers. These numbers represent a series of payments (negative values) and income (positive values) occurring at regular periods.
    - `Arg2 As Double` (required): Finance_rate - the interest rate you pay on the money used in the cash flows.
    - `Arg3 As Double` (required): Reinvest_rate - the interest rate you receive on the cash flows as you reinvest them.
- `Irr(Arg1 As Variant, [Arg2 As Variant]) As Double`  
  Returns the internal rate of return for a series of cash flows represented by the numbers in values. These cash flows don't have to be even, as they would be for an annuity. However, the cash flows must occur at regular intervals, such as monthly or annually. The internal rate of return is the interest rate received for an investment consisting of payments (negative values) and income (positive values) that occur at regular periods.
    - `Arg1 As Variant` (required): Values - an array or a reference to cells that contain numbers for which you want to calculate the internal rate of return.
    - `Arg2 As Variant` (optional): Guess - a number that you guess is close to the result of Irr.
- `Match(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As Double`  
  Returns the relative position of an item in an array that matches a specified value in a specified order. Use Match instead of one of the Lookup functions when you need the position of an item in a range instead of the item itself.
    - `Arg1 As Variant` (required): Lookup_value: the value that you use to find the value that you want in a table.
    - `Arg2 As Variant` (required): Lookup_array: a contiguous range of cells containing possible lookup values. Lookup_array must be an array or an array reference.
    - `Arg3 As Variant` (optional): Match_type: the number -1, 0, or 1. Match_type specifies how Microsoft Excel matches lookup_value with values in lookup_array.
- `Weekday(Arg1 As Variant, [Arg2 As Variant]) As Double`  
  Returns the day of the week corresponding to a date. The day is given as an integer, ranging from 1 (Sunday) to 7 (Saturday), by default.
    - `Arg1 As Variant` (required): Serial_number - a sequential number that represents the date of the day you are trying to find. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
    - `Arg2 As Variant` (optional): Return_type - a number that determines the type of return value.
- `Search(Arg1 As String, Arg2 As String, [Arg3 As Variant]) As Double`  
  Search and SearchB locate one text string within a second text string, and return the number of the starting position of the first text string from the first character of the second text string.
    - `Arg1 As String` (required): Find_text - the text that you want to find.
    - `Arg2 As String` (required): Within_text - the text in which you want to search for find_text.
    - `Arg3 As Variant` (optional): Start_num - the character number in within_text at which you want to start searching.
- `Transpose(Arg1 As Variant) As Variant`  
  Returns a vertical range of cells as a horizontal range, or vice versa. Transpose must be entered as an array formula in a range that has the same number of rows and columns, respectively, as an array has columns and rows. Use Transpose to shift the vertical and horizontal orientation of an array on a worksheet.
    - `Arg1 As Variant` (required): Array - an array or range of cells on a worksheet that you want to transpose. The transpose of an array is created by using the first row of the array as the first column of the new array, the second row of the array as the second column of the new array, and so on.
- `Atan2(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the arctangent, or inverse tangent, of the specified x- and y-coordinates. The arctangent is the angle from the x-axis to a line containing the origin (0, 0) and a point with coordinates (x_num, y_num). The angle is given in radians between -pi and pi, excluding -pi.
    - `Arg1 As Double` (required): The x-coordinate of the point.
    - `Arg2 As Double` (required): The y-coordinate of the point.
- `Asin(Arg1 As Double) As Double`  
  Returns the arcsine, or inverse sine, of a number. The arcsine is the angle whose sine is _Arg1_. The returned angle is given in radians in the range -pi/2 to pi/2.
    - `Arg1 As Double` (required): The sine of the angle that you want; must be from -1 to 1.
- `Acos(Arg1 As Double) As Double`  
  Returns the arccosine, or inverse cosine, of a number. The arccosine is the angle whose cosine is _Arg1_. The returned angle is given in radians in the range 0 (zero) to pi.
    - `Arg1 As Double` (required): The cosine of the angle you want, and must be from -1 to 1.
- `Choose(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Variant`  
  Uses _Arg1_ as the index to return a value from the list of value arguments.
    - `Arg1 As Variant` (required): Specifies which value argument is selected. _Arg1_ must be a number between 1 and 29, or a formula or reference to a cell containing a number between 1 and 29.
- `HLookup(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant]) As Variant`  
  Searches for a value in the top row of a table or an array of values, and then returns a value in the same column from a row that you specify in the table or array. Use HLookup when your comparison values are located in a row across the top of a table of data, and you want to look down a specified number of rows. Use VLookup when your comparison values are located in a column to the left of the data that you want to find.
    - `Arg1 As Variant` (required): Lookup_value - the value to be found in the first row of the table. Lookup_value can be a value, a reference, or a text string.
    - `Arg2 As Variant` (required): Table_array - a table of information in which data is looked up. Use a reference to a range or a range name.
    - `Arg3 As Variant` (required): Row_index_num - the row number in table_array from which the matching value will be returned. A row_index_num of 1 returns the first row value in table_array, a row_index_num of 2 returns the second row value in table_array, and so on. If row_index_num is less than 1, HLookup returns the #VALUE! error value; if row_index_num is greater than the number of rows on table_array, HLookup returns the #REF! error value.
    - `Arg4 As Variant` (optional): Range_lookup - a logical value that specifies whether you want HLookup to find an exact match or an approximate match. If True or omitted, an approximate match is returned. In other words, if an exact match is not found, the next largest value that is less than lookup_value is returned. If False, HLookup will find an exact match. If one is not found, the error value #N/A is returned.
- `VLookup(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant]) As Variant`  
  Searches for a value in the first column of a table array and returns a value in the same row from another column in the table array.
    - `Arg1 As Variant` (required): Lookup_value - the value to search in the first column of the table array. Lookup_value can be a value or a reference. If lookup_value is smaller than the smallest value in the first column of table_array, VLookup returns the #N/A error value.
    - `Arg2 As Variant` (required): Table_array - two or more columns of data. Use a reference to a range or a range name. The values in the first column of table_array are the values searched by lookup_value. These values can be text, numbers, or logical values. Uppercase and lowercase text are equivalent.
    - `Arg3 As Variant` (required): Col_index_num - the column number in table_array from which the matching value must be returned. A col_index_num of 1 returns the value in the first column in table_array; a col_index_num of 2 returns the value in the second column in table_array, and so on.
    - `Arg4 As Variant` (optional): Range_lookup - a logical value that specifies whether you want the VLookup method to find an exact match or an approximate match.
- `Log(Arg1 As Double, [Arg2 As Variant]) As Double`  
  Returns the logarithm of a number to the base that you specify.
    - `Arg1 As Double` (required): Number - the positive real number for which you want the logarithm.
    - `Arg2 As Variant` (optional): Base - the base of the logarithm. If base is omitted, it is assumed to be 10.
- `Proper(Arg1 As String) As String`  
  Capitalizes the first letter in a text string and any other letters in text that follow any character other than a letter. Converts all other letters to lowercase letters.
    - `Arg1 As String` (required): Text - text enclosed in quotation marks, a formula that returns text, or a reference to a cell containing the text that you want to partially capitalize.
- `Trim(Arg1 As String) As String`  
  Removes all spaces from text except for single spaces between words. Use Trim on text that you have received from another application that may have irregular spacing.
    - `Arg1 As String` (required): Text - the text from which you want spaces removed.
- `Replace(Arg1 As String, Arg2 As Double, Arg3 As Double, Arg4 As String) As String`  
  Replaces part of a text string, based on the number of characters that you specify, with a different text string.
    - `Arg1 As String` (required): Text in which you want to replace some characters.
    - `Arg2 As Double` (required): The position of the character in _Arg1_ that you want to replace with _Arg4_.
    - `Arg3 As Double` (required): The number of characters in _Arg1_ that you want the Replace method to replace with _Arg4_.
    - `Arg4 As String` (required): Text that will replace characters in _Arg1_.
- `Substitute(Arg1 As String, Arg2 As String, Arg3 As String, [Arg4 As Variant]) As String`  
  Substitutes new_text for old_text in a text string. Use Substitute when you want to replace specific text in a text string; use Replace when you want to replace any text that occurs in a specific location in a text string.
    - `Arg1 As String` (required): Text - the text or the reference to a cell containing text for which you want to substitute characters.
    - `Arg2 As String` (required): Old_text - the text that you want to replace.
    - `Arg3 As String` (required): New_text - the text that you want to replace old_text with.
    - `Arg4 As Variant` (optional): Instance_num - specifies which occurrence of old_text you want to replace with new_text. If you specify instance_num, only that instance of old_text is replaced. Otherwise, every occurrence of old_text in text is changed to new_text.
- `Find(Arg1 As String, Arg2 As String, [Arg3 As Variant]) As Double`  
  Finds specific information on a worksheet.
    - `Arg1 As String` (required): The name of the worksheet.
    - `Arg2 As String` (required): The name of the range.
    - `Arg3 As Variant` (optional): The name of an argument to refine the search.
- `IsErr(Arg1 As Variant) As Boolean`  
  Checks the type of value and returns True or False depending on whether the value refers to any error value except #N/A.
    - `Arg1 As Variant` (required): Value - the value that you want tested. Value can be a blank (empty cell), error, logical, text, number, or reference value, or a name referring to any of these, that you want to test.
- `IsText(Arg1 As Variant) As Boolean`  
  Checks the type of value and returns True or False depending on whether the value refers to text.
    - `Arg1 As Variant` (required): Value - the value that you want tested. Value can be a blank (empty cell), error, logical, text, number, or reference value, or a name referring to any of these, that you want to test.
- `IsNumber(Arg1 As Variant) As Boolean`  
  Checks the type of value and returns True or False depending on whether the value refers to a number.
    - `Arg1 As Variant` (required): Value - the value that you want tested. Value can be a blank (empty cell), error, logical, text, number, or reference value, or a name referring to any of these, that you want to test.
- `Sln(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the straight-line depreciation of an asset for one period.
    - `Arg1 As Double` (required): Cost - the initial cost of the asset.
    - `Arg2 As Double` (required): Salvage - the value at the end of the depreciation (sometimes called the salvage value of the asset).
    - `Arg3 As Double` (required): Life - the number of periods over which the asset is depreciated (sometimes called the useful life of the asset).
- `Syd(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Double) As Double`  
  Returns the sum-of-years' digits depreciation of an asset for a specified period.
    - `Arg1 As Double` (required): Cost - the initial cost of the asset.
    - `Arg2 As Double` (required): Salvage - the value at the end of the depreciation (sometimes called the salvage value of the asset).
    - `Arg3 As Double` (required): Life - the number of periods over which the asset is depreciated (sometimes called the useful life of the asset).
    - `Arg4 As Double` (required): Per - the period and must use the same units as life.
- `Ddb(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Double, [Arg5 As Variant]) As Double`  
  Returns the depreciation of an asset for a specified period by using the double-declining balance method or some other method that you specify.
    - `Arg1 As Double` (required): Cost - the initial cost of the asset.
    - `Arg2 As Double` (required): Salvage - the value at the end of the depreciation (sometimes called the salvage value of the asset). This value can be 0.
    - `Arg3 As Double` (required): Life - the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
    - `Arg4 As Double` (required): Period - the period for which you want to calculate the depreciation. Period must use the same units as life.
    - `Arg5 As Variant` (optional): Factor - the rate at which the balance declines. If factor is omitted, it is assumed to be 2 (the double-declining balance method).
- `Clean(Arg1 As String) As String`  
  Removes all nonprintable characters from text.
    - `Arg1 As String` (required): Any worksheet information from which you want to remove nonprintable characters.
- `MDeterm(Arg1 As Variant) As Double`  
  Returns the matrix determinant of an array.
    - `Arg1 As Variant` (required): Array - a numeric array with an equal number of rows and columns.
- `MInverse(Arg1 As Variant) As Variant`  
  Returns the inverse matrix for the matrix stored in an array.
    - `Arg1 As Variant` (required): Array - a numeric array with an equal number of rows and columns.
- `MMult(Arg1 As Variant, Arg2 As Variant) As Variant`  
  Returns the matrix product of two arrays. The result is an array with the same number of rows as array1 and the same number of columns as array2.
- `Ipmt(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Double, [Arg5 As Variant], [Arg6 As Variant]) As Double`  
  Returns the interest payment for a given period for an investment based on periodic, constant payments and a constant interest rate.
    - `Arg1 As Double` (required): Rate - the interest rate per period.
    - `Arg2 As Double` (required): Per - the period for which you want to find the interest; must be in the range 1 to nper.
    - `Arg3 As Double` (required): Nper - the total number of payment periods in an annuity.
    - `Arg4 As Double` (required): Pv - the present value, or the lump-sum amount that a series of future payments is worth right now.
    - `Arg5 As Variant` (optional): Fv - the future value, or a cash balance that you want to attain after the last payment is made. If fv is omitted, it is assumed to be 0 (the future value of a loan, for example, is 0).
    - `Arg6 As Variant` (optional): Type - the number 0 or 1 and indicates when payments are due. If type is omitted, it is assumed to be 0.
- `Ppmt(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Double, [Arg5 As Variant], [Arg6 As Variant]) As Double`  
  Returns the payment on the principal for a given period for an investment based on periodic, constant payments and a constant interest rate.
    - `Arg1 As Double` (required): Rate - the interest rate per period.
    - `Arg2 As Double` (required): Per - the period and must be in the range 1 to nper.
    - `Arg3 As Double` (required): Nper - the total number of payment periods in an annuity.
    - `Arg4 As Double` (required): Pv - the present value&mdash;the total amount that a series of future payments is worth now.
    - `Arg5 As Variant` (optional): Fv - the future value, or a cash balance you want to attain after the last payment is made. If fv is omitted, it is assumed to be 0 (zero), that is, the future value of a loan is 0.
    - `Arg6 As Variant` (optional): Type - the number 0 or 1 and indicates when payments are due.
- `CountA(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Counts the number of cells that are not empty and the values within the list of arguments.
- `Product(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Multiplies all the numbers given as arguments and returns the product.
- `Fact(Arg1 As Double) As Double`  
  Returns the factorial of a number. The factorial of a number is equal to 1\2\3\...\ number.
    - `Arg1 As Double` (required): Number - the nonnegative number that you want the factorial of. If number is not an integer, it is truncated.
- `DProduct(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Double`  
  Multiplies the values in a column of a list or database that match conditions that you specify.
    - `Arg1 As Range` (required): Database is the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Field indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): Criteria is the range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `IsNonText(Arg1 As Variant) As Boolean`  
  Checks the type of value and returns True or False depending on whether the value refers to any item that is not text. (Note that this function returns True if value refers to a blank cell.)
    - `Arg1 As Variant` (required): Value - the value that you want tested. Value can be a blank (empty cell), error, logical, text, number, or reference value, or a name referring to any of these, that you want to test.
- `StDevP(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Calculates standard deviation based on the entire population given as arguments. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).
- `VarP(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Calculates variance based on the entire population.
- `DStDevP(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Double`  
  Calculates the standard deviation of a population based on the entire population by using the numbers in a column of a list or database that match conditions that you specify.
    - `Arg1 As Range` (required): Database - is the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Field - indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): Criteria - the range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `DVarP(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Double`  
  Calculates the variance of a population based on the entire population by using the numbers in a column of a list or database that match conditions that you specify.
    - `Arg1 As Range` (required): Database - the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Field - indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): Criteria - the range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `IsLogical(Arg1 As Variant) As Boolean`  
  Checks the type of value and returns True or False depending on whether the value refers to a logical value.
    - `Arg1 As Variant` (required): Value - the value that you want tested. Value can be a blank (empty cell), error, logical, text, number, or reference value, or a name referring to any of these, that you want to test.
- `DCountA(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Double`  
  Counts the nonblank cells in a column of a list or database that match conditions that you specify.
    - `Arg1 As Range` (required): Database - the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Field - indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): Criteria - the range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `USDollar(Arg1 As Double, Arg2 As Double) As String`  
  Converts a number to text format and applies a currency symbol. The name of the method (and the symbol that it applies) depends upon the language settings.
    - `Arg1 As Double` (required): A reference to a cell containing a number, or a formula that evaluates to a number.
    - `Arg2 As Double` (required): The number of digits to the right of the decimal point. If _Arg2_ is negative, the number is rounded to the left of the decimal point. If you omit decimals, it is assumed to be 2.
- `FindB(Arg1 As String, Arg2 As String, [Arg3 As Variant]) As Double`  
  Find and FindB locate one text string within a second text string, and return the number of the starting position of the first text string from the first character of the second text string.
    - `Arg1 As String` (required): Find_text - the text you want to find.
    - `Arg2 As String` (required): Within_text - the text containing the text that you want to find.
    - `Arg3 As Variant` (optional): Start_num - specifies the character at which to start the search. The first character in within_text is character number 1. If you omit start_num, it is assumed to be 1.
- `SearchB(Arg1 As String, Arg2 As String, [Arg3 As Variant]) As Double`  
  Search and SearchB locate one text string within a second text string, and return the number of the starting position of the first text string from the first character of the second text string.
    - `Arg1 As String` (required): Find_text - the text that you want to find.
    - `Arg2 As String` (required): Within_text - the text in which you want to search for find_text.
    - `Arg3 As Variant` (optional): Start_num - the character number in within_text at which you want to start searching.
- `ReplaceB(Arg1 As String, Arg2 As Double, Arg3 As Double, Arg4 As String) As String`  
  Replaces part of a text string, based on the number of bytes that you specify, with a different text string.
    - `Arg1 As String` (required): Old_text - text in which you want to replace some characters.
    - `Arg2 As Double` (required): Start_num - the position of the character in old_text that you want to replace with new_text.
    - `Arg3 As Double` (required): Num_chars - the number of characters in old_text that you want Replace to replace with new_text.
    - `Arg4 As String` (required): New_text - the text that will replace characters in old_text.
- `RoundUp(Arg1 As Double, Arg2 As Double) As Double`  
  Rounds a number up, away from 0 (zero).
    - `Arg1 As Double` (required): Number - any real number that you want rounded up.
    - `Arg2 As Double` (required): Num_digits - the number of digits to which you want to round number.
- `RoundDown(Arg1 As Double, Arg2 As Double) As Double`  
  Rounds a number down, toward 0 (zero).
    - `Arg1 As Double` (required): Number - any real number that you want rounded down.
    - `Arg2 As Double` (required): Num_digits - the number of digits to which you want to round number.
- `Rank(Arg1 As Double, Arg2 As Range, [Arg3 As Variant]) As Double`  
  Returns the rank of a number in a list of numbers. The rank of a number is its size relative to other values in a list. If you were to sort the list, the rank of the number would be its position.
    - `Arg1 As Double` (required): Number - the number whose rank you want to find.
    - `Arg2 As Range` (required): Ref - an array of, or a reference to, a list of numbers. Nonnumeric values in ref are ignored.
    - `Arg3 As Variant` (optional): Order - a number specifying how to rank number.
- `Days360(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As Double`  
  Returns the number of days between two dates based on a 360-day year (twelve 30-day months), which is used in some accounting calculations.
    - `Arg3 As Variant` (optional): A Boolean value that specifies whether to use the U.S. or European method in the calculation.
- `Vdb(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Double, Arg5 As Double, [Arg6 As Variant], [Arg7 As Variant]) As Double`  
  Returns the depreciation of an asset for any period that you specify, including partial periods, by using the double-declining balance method or some other method that you specify. Vdb stands for variable declining balance.
    - `Arg1 As Double` (required): Cost - the initial cost of the asset.
    - `Arg2 As Double` (required): Salvage - the value at the end of the depreciation (sometimes called the salvage value of the asset). This value can be 0.
    - `Arg3 As Double` (required): Life - the number of periods over which the asset is depreciated (sometimes called the useful life of the asset).
    - `Arg4 As Double` (required): Start_period - the starting period for which you want to calculate the depreciation. Start_period must use the same units as life.
    - `Arg5 As Double` (required): End_period - the ending period for which you want to calculate the depreciation. End_period must use the same units as life.
    - `Arg6 As Variant` (optional): Factor - the rate at which the balance declines. If factor is omitted, it is assumed to be 2 (the double-declining balance method). Change factor if you don't want to use the double-declining balance method. For a description of the double-declining balance method, see Ddb.
    - `Arg7 As Variant` (optional): No_switch - a logical value specifying whether to switch to straight-line depreciation when depreciation is greater than the declining balance calculation.
- `Median(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the median of the given numbers. The median is the number in the middle of a set of numbers.
- `SumProduct(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Multiplies corresponding components in the given arrays, and returns the sum of those products.
- `Sinh(Arg1 As Double) As Double`  
  Returns the hyperbolic sine of a number.
    - `Arg1 As Double` (required): Number is any real number.
- `Cosh(Arg1 As Double) As Double`  
  Returns the hyperbolic cosine of a number.
    - `Arg1 As Double` (required): Any real number for which you want to find the hyperbolic cosine.
- `Tanh(Arg1 As Double) As Double`  
  Returns the hyperbolic tangent of a number.
    - `Arg1 As Double` (required): Number - any real number.
- `Asinh(Arg1 As Double) As Double`  
  Returns the inverse hyperbolic sine of a number. The inverse hyperbolic sine is the value whose hyperbolic sine is _Arg1_, so Asinh(Sinh(number)) equals _Arg1_.
    - `Arg1 As Double` (required): Any real number.
- `Acosh(Arg1 As Double) As Double`  
  Returns the inverse hyperbolic cosine of a number. Number must be greater than or equal to 1. The inverse hyperbolic cosine is the value whose hyperbolic cosine is _Arg1_, so Acosh(Cosh(number)) equals _Arg1_.
    - `Arg1 As Double` (required): Any real number equal to or greater than 1.
- `Atanh(Arg1 As Double) As Double`  
  Returns the inverse hyperbolic tangent of a number. Number must be between -1 and 1 (excluding -1 and 1).
    - `Arg1 As Double` (required): Any real number between 1 and -1.
- `DGet(Arg1 As Range, Arg2 As Variant, Arg3 As Variant) As Variant`  
  Extracts a single value from a column of a list or database that matches conditions that you specify.
    - `Arg1 As Range` (required): Database - the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.
    - `Arg2 As Variant` (required): Field - indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.
    - `Arg3 As Variant` (required): Criteria - the range of cells that contains the conditions that you specify. Use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.
- `Db(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Double, [Arg5 As Variant]) As Double`  
  Returns the depreciation of an asset for a specified period using the fixed-declining balance method.
    - `Arg1 As Double` (required): Cost - the initial cost of the asset.
    - `Arg2 As Double` (required): Salvage - the value at the end of the depreciation (sometimes called the salvage value of the asset).
    - `Arg3 As Double` (required): Life - the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
    - `Arg4 As Double` (required): Period - the period for which you want to calculate the depreciation. Period must use the same units as life.
    - `Arg5 As Variant` (optional): Month - the number of months in the first year. If month is omitted, it is assumed to be 12.
- `Frequency(Arg1 As Variant, Arg2 As Variant) As Variant`  
  Calculates how often values occur within a range of values, and then returns a vertical array of numbers. For example, use Frequency to count the number of test scores that fall within ranges of scores. Because Frequency returns an array, it must be entered as an array formula.
    - `Arg1 As Variant` (required): Data_array - an array of or reference to a set of values for which you want to count frequencies. If data_array contains no values, Frequency returns an array of zeros.
    - `Arg2 As Variant` (required): Bins_array - an array of or reference to intervals into which you want to group the values in data_array. If bins_array contains no values, Frequency returns the number of elements in data_array.
- `AveDev(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the average of the absolute deviations of data points from their mean. AveDev is a measure of the variability in a data set.
- `BetaDist(Arg1 As Double, Arg2 As Double, Arg3 As Double, [Arg4 As Variant], [Arg5 As Variant]) As Double`  
  Returns the beta cumulative distribution function.
    - `Arg1 As Double` (required): The value between A and B at which to evaluate the function.
    - `Arg2 As Double` (required): A parameter of the distribution.
    - `Arg3 As Double` (required): A parameter of the distribution.
    - `Arg4 As Variant` (optional): An optional lower bound to the interval of x.
    - `Arg5 As Variant` (optional): An optional upper bound to the interval of x.
- `GammaLn(Arg1 As Double) As Double`  
  Returns the natural logarithm of the gamma function, (x).
    - `Arg1 As Double` (required): X - the value for which you want to calculate GammaLn.
- `BetaInv(Arg1 As Double, Arg2 As Double, Arg3 As Double, [Arg4 As Variant], [Arg5 As Variant]) As Double`  
  Returns the inverse of the cumulative distribution function for a specified beta distribution. That is, if probability = BetaDist(x,...), then BetaInv(probability,...) = x.
    - `Arg1 As Double` (required): A probability associated with the beta distribution.
    - `Arg2 As Double` (required): The Alpha parameter of the distribution.
    - `Arg3 As Double` (required): The Beta parameter the distribution.
    - `Arg4 As Variant` (optional): An optional lower bound to the interval of x.
    - `Arg5 As Variant` (optional): An optional upper bound to the interval of x.
- `BinomDist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean) As Double`  
  Returns the individual term binomial distribution probability.
    - `Arg1 As Double` (required): The number of successes in trials.
    - `Arg2 As Double` (required): The number of independent trials.
    - `Arg3 As Double` (required): The probability of success on each trial.
    - `Arg4 As Boolean` (required): A logical value that determines the form of the function. If cumulative is True, BinomDist returns the cumulative distribution function, which is the probability that there are at most number_s successes; if False, it returns the probability mass function, which is the probability that there are number_s successes.
- `ChiDist(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the one-tailed probability of the chi-squared distribution.
    - `Arg1 As Double` (required): The value at which you want to evaluate the distribution.
    - `Arg2 As Double` (required): The number of degrees of freedom.
- `ChiInv(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the inverse of the one-tailed probability of the chi-squared distribution.
    - `Arg1 As Double` (required): A probability associated with the chi-squared distribution.
    - `Arg2 As Double` (required): The number of degrees of freedom.
- `Combin(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the number of combinations for a given number of items. Use Combin to determine the total possible number of groups for a given number of items.
    - `Arg1 As Double` (required): The number of items.
    - `Arg2 As Double` (required): The number of items in each combination.
- `Confidence(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns a value that you can use to construct a confidence interval for a population mean.
    - `Arg1 As Double` (required): The significance level used to compute the confidence level. The confidence level equals 100*(1 - alpha)%, or in other words, an alpha of 0.05 indicates a 95 percent confidence level.
    - `Arg2 As Double` (required): The population standard deviation for the data range; is assumed to be known.
    - `Arg3 As Double` (required): The sample size.
- `CritBinom(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.
    - `Arg1 As Double` (required): The number of Bernoulli trials.
    - `Arg2 As Double` (required): The probability of a success on each trial.
    - `Arg3 As Double` (required): The criterion value.
- `Even(Arg1 As Double) As Double`  
  Returns number rounded up to the nearest even integer. Use this function for processing items that come in twos. For example, a packing crate accepts rows of one or two items. The crate is full when the number of items, rounded up to the nearest two, matches the crate's capacity.
    - `Arg1 As Double` (required): Number - the value to round.
- `ExponDist(Arg1 As Double, Arg2 As Double, Arg3 As Boolean) As Double`  
  Returns the exponential distribution. Use ExponDist to model the time between events, such as how long an automated bank teller takes to deliver cash. For example, you can use ExponDist to determine the probability that the process takes at most 1 minute.
    - `Arg1 As Double` (required): X - the value of the function.
    - `Arg2 As Double` (required): Lambda - the parameter value.
    - `Arg3 As Boolean` (required): Cumulative - a logical value that indicates which form of the exponential function to provide. If cumulative is True, ExponDist returns the cumulative distribution function; if False, it returns the probability density function.
- `FDist(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the F probability distribution. Use this function to determine whether two data sets have different degrees of diversity. For example, you can examine the test scores of men and women entering high school and determine if the variability in the females is different from that found in the males.
    - `Arg1 As Double` (required): X - the value at which to evaluate the function.
    - `Arg2 As Double` (required): Degrees_freedom1 - the numerator degrees of freedom.
    - `Arg3 As Double` (required): Degrees_freedom2 - the denominator degrees of freedom.
- `FInv(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the inverse of the F probability distribution. If p = FDIST(x,...), then FINV(p,...) = x.
    - `Arg1 As Double` (required): Probability - a probability associated with the F cumulative distribution.
    - `Arg2 As Double` (required): Degrees_freedom1 - the numerator degrees of freedom.
    - `Arg3 As Double` (required): Degrees_freedom2 - is the denominator degrees of freedom.
- `Fisher(Arg1 As Double) As Double`  
  Returns the Fisher transformation at x. This transformation produces a function that is normally distributed rather than skewed. Use this function to perform hypothesis testing on the correlation coefficient.
    - `Arg1 As Double` (required): x - a numeric value for which you want the transformation.
- `FisherInv(Arg1 As Double) As Double`  
  Returns the inverse of the Fisher transformation. Use this transformation when analyzing correlations between ranges or arrays of data. If y = FISHER(x), then FISHERINV(y) = x.
    - `Arg1 As Double` (required): y - the value for which you want to perform the inverse of the transformation.
- `Floor(Arg1 As Double, Arg2 As Double) As Double`  
  Rounds number down, toward zero, to the nearest multiple of significance.
    - `Arg1 As Double` (required): Number - the numeric value that you want to round.
    - `Arg2 As Double` (required): Significance - the multiple to which you want to round.
- `GammaDist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean) As Double`  
  Returns the gamma distribution. Use this function to study variables that may have a skewed distribution. The gamma distribution is commonly used in queuing analysis.
    - `Arg1 As Double` (required): X - the value at which you want to evaluate the distribution.
    - `Arg2 As Double` (required): Alpha - a parameter to the distribution.
    - `Arg3 As Double` (required): Beta - a parameter to the distribution. If beta = 1, GammaDist returns the standard gamma distribution.
    - `Arg4 As Boolean` (required): Cumulative - a logical value that determines the form of the function. If cumulative is True, GammaDist returns the cumulative distribution function; if False, it returns the probability density function.
- `GammaInv(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the inverse of the gamma cumulative distribution. If p = GAMMADIST(x,...), then GAMMAINV(p,...) = x.
    - `Arg1 As Double` (required): Probability - the probability associated with the gamma distribution.
    - `Arg2 As Double` (required): Alpha - a parameter to the distribution.
    - `Arg3 As Double` (required): Beta - a parameter to the distribution. If beta = 1, GammaInv returns the standard gamma distribution.
- `Ceiling(Arg1 As Double, Arg2 As Double) As Double`  
  Returns number rounded up, away from zero, to the nearest multiple of significance.
    - `Arg1 As Double` (required): Number - the value that you want to round.
    - `Arg2 As Double` (required): Significance - the multiple to which you want to round.
- `HypGeomDist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Double) As Double`  
  Returns the hypergeometric distribution. HypGeomDist returns the probability of a given number of sample successes, given the sample size, population successes, and population size. Use HypGeomDist for problems with a finite population, where each observation is either a success or a failure, and where each subset of a given size is chosen with equal likelihood.
    - `Arg1 As Double` (required): Sample_s - the number of successes in the sample.
    - `Arg2 As Double` (required): Number_sample - the size of the sample.
    - `Arg3 As Double` (required): Population_s - the number of successes in the population.
    - `Arg4 As Double` (required): Number_population - the population size.
- `LogNormDist(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the cumulative lognormal distribution of x, where ln(x) is normally distributed with parameters mean and standard_dev. Use this function to analyze data that has been logarithmically transformed.
    - `Arg1 As Double` (required): X - the value at which to evaluate the function.
    - `Arg2 As Double` (required): Mean - the mean of ln(x).
    - `Arg3 As Double` (required): Standard_dev - the standard deviation of ln(x).
- `LogInv(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Use the lognormal distribution to analyze logarithmically transformed data.
    - `Arg1 As Double` (required): Probability - a probability associated with the lognormal distribution.
    - `Arg2 As Double` (required): Mean - the mean of ln(x).
    - `Arg3 As Double` (required): Standard_dev - the standard deviation of ln(x).
- `NegBinomDist(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the negative binomial distribution. NegBinomDist returns the probability that there will be number_f failures before the number_s-th success, when the constant probability of a success is probability_s. This function is similar to the binomial distribution, except that the number of successes is fixed, and the number of trials is variable. Like the binomial, trials are assumed to be independent.
    - `Arg1 As Double` (required): Number_f - the number of failures.
    - `Arg2 As Double` (required): Number_s - the threshold number of successes.
    - `Arg3 As Double` (required): Probability_s - the probability of a success.
- `NormDist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean) As Double`  
  Returns the normal distribution for the specified mean and standard deviation. This function has a very wide range of applications in statistics, including hypothesis testing.
    - `Arg1 As Double` (required): X - the value for which you want the distribution.
    - `Arg2 As Double` (required): Mean - the arithmetic mean of the distribution.
    - `Arg3 As Double` (required): Standard_dev - the standard deviation of the distribution.
    - `Arg4 As Boolean` (required): Cumulative - a logical value that determines the form of the function. If cumulative is True, NormDist returns the cumulative distribution function; if False, it returns the probability mass function.
- `NormSDist(Arg1 As Double) As Double`  
  Returns the standard normal cumulative distribution function. The distribution has a mean of 0 (zero) and a standard deviation of one. Use this function in place of a table of standard normal curve areas.
    - `Arg1 As Double` (required): Z - the value for which you want the distribution.
- `NormInv(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.
    - `Arg1 As Double` (required): Probability - a probability corresponding to the normal distribution.
    - `Arg2 As Double` (required): Mean - the arithmetic mean of the distribution.
    - `Arg3 As Double` (required): Standard_dev - the standard deviation of the distribution.
- `NormSInv(Arg1 As Double) As Double`  
  Returns the inverse of the standard normal cumulative distribution. The distribution has a mean of zero and a standard deviation of one.
    - `Arg1 As Double` (required): Probability - a probability corresponding to the normal distribution.
- `Standardize(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns a normalized value from a distribution characterized by mean and standard_dev.
    - `Arg1 As Double` (required): X - the value that you want to normalize.
    - `Arg2 As Double` (required): Mean - the arithmetic mean of the distribution.
    - `Arg3 As Double` (required): Standard_dev - the standard deviation of the distribution.
- `Odd(Arg1 As Double) As Double`  
  Returns number rounded up to the nearest odd integer.
    - `Arg1 As Double` (required): Number - the value to round.
- `Permut(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the number of permutations for a given number of objects that can be selected from number objects. A permutation is any set or subset of objects or events where internal order is significant. Permutations are different from combinations, for which the internal order is not significant. Use this function for lottery-style probability calculations.
    - `Arg1 As Double` (required): Number - an integer that describes the number of objects.
    - `Arg2 As Double` (required): Number_chosen - an integer that describes the number of objects in each permutation.
- `Poisson(Arg1 As Double, Arg2 As Double, Arg3 As Boolean) As Double`  
  Returns the Poisson distribution. A common application of the Poisson distribution is predicting the number of events over a specific time, such as the number of cars arriving at a toll plaza in 1 minute.
    - `Arg1 As Double` (required): X - the number of events.
    - `Arg2 As Double` (required): Mean - the expected numeric value.
    - `Arg3 As Boolean` (required): Cumulative - a logical value that determines the form of the probability distribution returned. If cumulative is True, Poisson returns the cumulative Poisson probability that the number of random events occurring will be between zero and x inclusive; if False, it returns the Poisson probability mass function that the number of events occurring will be exactly x.
- `TDist(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the Percentage Points (probability) for the Student t-distribution where a numeric value (x) is a calculated value of t for which the Percentage Points are to be computed. The t-distribution is used in the hypothesis testing of small sample data sets. Use this function in place of a table of critical values for the t-distribution.
    - `Arg1 As Double` (required): X - the numeric value at which to evaluate the distribution.
    - `Arg2 As Double` (required): Degrees_freedom - an integer indicating the number of degrees of freedom.
    - `Arg3 As Double` (required): Tails - specifies the number of distribution tails to return. If tails = 1, TDist returns the one-tailed distribution. If tails = 2, TDist returns the two-tailed distribution.
- `Weibull(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean) As Double`  
  Returns the Weibull distribution. Use this distribution in reliability analysis, such as calculating a device's mean time to failure.
    - `Arg1 As Double` (required): X - the value at which to evaluate the function.
    - `Arg2 As Double` (required): Alpha - a parameter to the distribution.
    - `Arg3 As Double` (required): Beta - a parameter to the distribution.
    - `Arg4 As Boolean` (required): Cumulative - determines the form of the function.
- `SumXMY2(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the sum of squares of differences of corresponding values in two arrays.
    - `Arg1 As Variant` (required): Array_x - the first array or range of values.
    - `Arg2 As Variant` (required): Array_y - the second array or range of values.
- `SumX2MY2(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the sum of the difference of squares of corresponding values in two arrays.
    - `Arg1 As Variant` (required): Array_x - the first array or range of values.
    - `Arg2 As Variant` (required): Array_y - the second array or range of values.
- `SumX2PY2(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the sum of the sum of squares of corresponding values in two arrays. The sum of the sum of squares is a common term in many statistical calculations.
    - `Arg1 As Variant` (required): Array_x - the first array or range of values.
    - `Arg2 As Variant` (required): Array_y - the second array or range of values.
- `ChiTest(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the test for independence.
    - `Arg1 As Variant` (required): The range of data that contains observations to test against expected values.
    - `Arg2 As Variant` (required): The range of data that contains the ratio of the product of row totals and column totals to the grand total.
- `Correl(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the correlation coefficient of the _Arg1_ and _Arg2_ cell ranges.
    - `Arg1 As Variant` (required): A cell range of values.
    - `Arg2 As Variant` (required): A second cell range of values.
- `Covar(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns covariance, the average of the products of deviations for each data point pair.
    - `Arg1 As Variant` (required): The first cell range of integers.
    - `Arg2 As Variant` (required): The second cell range of integers.
- `FTest(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the result of an F-test. An F-test returns the two-tailed probability that the variances in array1 and array2 are not significantly different. Use this function to determine whether two samples have different variances. For example, given test scores from public and private schools, you can test whether these schools have different levels of test score diversity.
    - `Arg1 As Variant` (required): Array1 - the first array or range of data.
    - `Arg2 As Variant` (required): Array2 - the second array or range of data.
- `Intercept(Arg1 As Variant, Arg2 As Variant) As Double`  
  Calculates the point at which a line will intersect the y-axis by using existing x-values and y-values. The intercept point is based on a best-fit regression line plotted through the known x-values and known y-values.
    - `Arg1 As Variant` (required): Known_y's - the dependent set of observations or data.
    - `Arg2 As Variant` (required): Known_x's - the independent set of observations or data.
- `Pearson(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the Pearson product moment correlation coefficient, r, a dimensionless index that ranges from -1.0 to 1.0 inclusive and reflects the extent of a linear relationship between two data sets.
    - `Arg1 As Variant` (required): Array1 - a set of independent values.
    - `Arg2 As Variant` (required): Array2 - a set of dependent values.
- `RSq(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the square of the Pearson product moment correlation coefficient through data points in known_y's and known_x's. For more information, see Pearson. The r-squared value can be interpreted as the proportion of the variance in y attributable to the variance in x.
    - `Arg1 As Variant` (required): Known_y's - an array or range of data points.
    - `Arg2 As Variant` (required): Known_x's - an array or range of data points.
- `StEyx(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the standard error of the predicted y-value for each x in the regression. The standard error is a measure of the amount of error in the prediction of y for an individual x.
    - `Arg1 As Variant` (required): Known_y's - an array or range of dependent data points.
    - `Arg2 As Variant` (required): Known_x's - an array or range of independent data points.
- `Slope(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the slope of the linear regression line through data points in known_y's and known_x's. The slope is the vertical distance divided by the horizontal distance between any two points on the line, which is the rate of change along the regression line.
    - `Arg1 As Variant` (required): Known_y's - an array or cell range of numeric dependent data points.
    - `Arg2 As Variant` (required): Known_x's - the set of independent data points.
- `TTest(Arg1 As Variant, Arg2 As Variant, Arg3 As Double, Arg4 As Double) As Double`  
  Returns the probability associated with a Student's t-Test. Use TTest to determine whether two samples are likely to have come from the same two underlying populations that have the same mean.
    - `Arg1 As Variant` (required): Array1 - the first data set.
    - `Arg2 As Variant` (required): Array2 - the second data set.
    - `Arg3 As Double` (required): Tails - specifies the number of distribution tails. If tails = 1, TTest uses the one-tailed distribution. If tails = 2, TTest uses the two-tailed distribution.
    - `Arg4 As Double` (required): Type - the kind of t-Test to perform.
- `Prob(Arg1 As Variant, Arg2 As Variant, Arg3 As Double, [Arg4 As Variant]) As Double`  
  Returns the probability that values in a range are between two limits. If upper_limit is not supplied, returns the probability that values in x_range are equal to lower_limit.
    - `Arg1 As Variant` (required): X_range - the range of numeric values of x with which there are associated probabilities.
    - `Arg2 As Variant` (required): Prob_range - a set of probabilities associated with values in x_range.
    - `Arg3 As Double` (required): Lower_limit - the lower bound on the value for which you want a probability.
    - `Arg4 As Variant` (optional): Upper_limit - the optional upper bound on the value for which you want a probability.
- `DevSq(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the sum of squares of deviations of data points from their sample mean.
- `GeoMean(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the geometric mean of an array or range of positive data. For example, you can use GeoMean to calculate average growth rate given compound interest with variable rates.
- `HarMean(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the harmonic mean of a data set. The harmonic mean is the reciprocal of the arithmetic mean of reciprocals.
- `SumSq(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the sum of the squares of the arguments.
    - `Arg1 As Variant` (required): Number1, number2... - 1 to 30 arguments for which you want the sum of the squares. You can also use a single array or a reference to an array instead of arguments separated by commas.
- `Kurt(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the kurtosis of a data set. Kurtosis characterizes the relative peakedness or flatness of a distribution compared with the normal distribution. Positive kurtosis indicates a relatively peaked distribution. Negative kurtosis indicates a relatively flat distribution.
- `Skew(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the skewness of a distribution. Skewness characterizes the degree of asymmetry of a distribution around its mean.
- `ZTest(Arg1 As Variant, Arg2 As Double, [Arg3 As Variant]) As Double`  
  Returns the one-tailed probability-value of a z-test. For a given hypothesized population mean, ZTest returns the probability that the sample mean would be greater than the average of observations in the data set (_array_); that is, the observed sample mean.
    - `Arg1 As Variant` (required): _Array_ is the array or range of data against which to test the hypothesized population mean.
    - `Arg2 As Double` (required): The value to test.
    - `Arg3 As Variant` (optional): _Sigma_ is the population (known) standard deviation. If omitted, the sample standard deviation is used.
- `Large(Arg1 As Variant, Arg2 As Double) As Double`  
  Returns the k-th largest value in a data set. Use this function to select a value based on its relative standing. For example, you can use Large to return the highest, runner-up, or third-place score.
    - `Arg1 As Variant` (required): Array - the array or range of data for which you want to determine the k-th largest value.
    - `Arg2 As Double` (required): K - the position (from the largest) in the array or cell range of data to return.
- `Small(Arg1 As Variant, Arg2 As Double) As Double`  
  Returns the k-th smallest value in a data set. Use this function to return values with a particular relative standing in a data set.
    - `Arg1 As Variant` (required): Array - an array or range of numerical data for which you want to determine the k-th smallest value.
    - `Arg2 As Double` (required): K - the position (from the smallest) in the array or range of data to return.
- `Quartile(Arg1 As Variant, Arg2 As Double) As Double`  
  Returns the quartile of a data set. Quartiles often are used in sales and survey data to divide populations into groups. For example, you can use Quartile to find the top 25 percent of incomes in a population.
    - `Arg1 As Variant` (required): Array - the array or cell range of numeric values for which you want the quartile value.
    - `Arg2 As Double` (required): Quart - indicates which value to return.
- `Percentile(Arg1 As Variant, Arg2 As Double) As Double`  
  Returns the k-th percentile of values in a range. Use this function to establish a threshold of acceptance. For example, you can decide to examine candidates who score above the 90th percentile.
    - `Arg1 As Variant` (required): Array - the array or range of data that defines relative standing.
    - `Arg2 As Double` (required): K - the percentile value in the range 0..1, inclusive.
- `PercentRank(Arg1 As Variant, Arg2 As Double, [Arg3 As Variant]) As Double`  
  Returns the rank of a value in a data set as a percentage of the data set. This function can be used to evaluate the relative standing of a value within a data set. For example, you can use PercentRank to evaluate the standing of an aptitude test score among all scores for the test.
    - `Arg1 As Variant` (required): Array - the array or range of data with numeric values that defines relative standing.
    - `Arg2 As Double` (required): X - the value for which you want to know the rank.
    - `Arg3 As Variant` (optional): Significance - an optional value that identifies the number of significant digits for the returned percentage value. If omitted, PercentRank uses three digits (0.xxx).
- `Mode(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the most frequently occurring, or repetitive, value in an array or range of data.
- `TrimMean(Arg1 As Variant, Arg2 As Double) As Double`  
  Returns the mean of the interior of a data set. TrimMean calculates the mean taken by excluding a percentage of data points from the top and bottom tails of a data set. Use this function when you wish to exclude outlying data from your analysis.
    - `Arg1 As Variant` (required): Array - the array or range of values to trim and average.
    - `Arg2 As Double` (required): Percent - the fractional number of data points to exclude from the calculation. For example, if percent = 0.2, 4 points are trimmed from a data set of 20 points (20 x 0.2): 2 from the top and 2 from the bottom of the set.
- `TInv(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the t-value of the Student's t-distribution as a function of the probability and the degrees of freedom.
    - `Arg1 As Double` (required): Probability - the probability associated with the two-tailed Student's t-distribution.
    - `Arg2 As Double` (required): Degrees_freedom - the number of degrees of freedom with which to characterize the distribution.
- `Power(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the result of a number raised to a power.
    - `Arg1 As Double` (required): Number - the base number. It can be any real number.
    - `Arg2 As Double` (required): Power - the exponent to which the base number is raised.
- `Radians(Arg1 As Double) As Double`  
  Converts degrees to radians.
    - `Arg1 As Double` (required): Angle - an angle in degrees that you want to convert.
- `Degrees(Arg1 As Double) As Double`  
  Converts radians into degrees.
    - `Arg1 As Double` (required): Angle - the angle in radians that you want to convert.
- `Subtotal(Arg1 As Double, Arg2 As Range, [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Creates subtotals.
    - `Arg1 As Double` (required): A number indicating the aggregation function to be used by the method.
    - `Arg2 As Range` (required): The first Range object for which a subtotal is to be calculated.
- `SumIf(Arg1 As Range, Arg2 As Variant, [Arg3 As Variant]) As Double`  
  Adds the cells specified by a given criteria.
    - `Arg1 As Range` (required): Range - the range of cells that you want evaluated by criteria.
    - `Arg2 As Variant` (required): Criteria - the criteria in the form of a number, expression, or text that defines which cells will be added. For example, criteria can be expressed as 32, "32", ">32", or "apples".
    - `Arg3 As Variant` (optional): Sum_range - the actual cells to add if their corresponding cells in range match criteria. If sum_range is omitted, the cells in range are both evaluated by criteria and added if they match criteria.
- `CountIf(Arg1 As Range, Arg2 As Variant) As Double`  
  Counts the number of cells within a range that meet the given criteria.
    - `Arg1 As Range` (required): The range of cells from which you want to count cells.
    - `Arg2 As Variant` (required): The criteria in the form of a number, expression, cell reference, or text that defines which cells will be counted. For example, criteria can be expressed as 32, "32", ">32", "apples", or B4.
- `CountBlank(Arg1 As Range) As Double`  
  Counts empty cells in a specified range of cells.
    - `Arg1 As Range` (required): The range from which you want to count the blank cells.
- `Ispmt(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Double) As Double`  
  Calculates the interest paid during a specific period of an investment. This function is provided for compatibility with Lotus 1-2-3.
    - `Arg1 As Double` (required): Rate - the interest rate for the investment.
    - `Arg2 As Double` (required): Per - the period for which you want to find the interest; must be between 1 and nper.
    - `Arg3 As Double` (required): Nper - the total number of payment periods for the investment.
    - `Arg4 As Double` (required): Pv - the present value of the investment. For a loan, pv is the loan amount.
- `Roman(Arg1 As Double, [Arg2 As Variant]) As String`  
  Converts an arabic numeral to roman, as text.
    - `Arg1 As Double` (required): Number - the Arabic numeral that you want converted.
    - `Arg2 As Variant` (optional): Form - a number specifying the type of roman numeral you want. The roman numeral style ranges from Classic to Simplified, becoming more concise as the value of form increases.
- `Asc(Arg1 As String) As String`  
  For double-byte character set (DBCS) languages, changes full-width (double-byte) characters to half-width (single-byte) characters.
    - `Arg1 As String` (required): The text or a reference to a cell that contains the text that you want to change. If the text does not contain any full-width letters, the text is not changed.
- `Dbcs(Arg1 As String) As String`  
  Converts half-width (single-byte) letters within a character string to full-width (double-byte) characters. The name of the function (and the characters that it converts) depends upon the language settings. Read/write String.
    - `Arg1 As String` (required): The text or a reference to a cell that contains the text that you want to change.
- `Phonetic(Arg1 As Range) As String`  
  Extracts the phonetic (furigana) characters from a text string.
    - `Arg1 As Range` (required): Reference - a text string or a reference to a single cell or a range of cells that contain a furigana text string.
- `BahtText(Arg1 As Double) As String`  
  Converts a number to Thai text and adds a suffix of Baht.
    - `Arg1 As Double` (required): A number that you want to convert to text, or a reference to a cell containing a number, or a formula that evaluates to a number.
- `RTD(progID As Variant, server As Variant, topic1 As Variant, [topic2 As Variant], [topic3 As Variant], [topic4 As Variant], [topic5 As Variant], [topic6 As Variant], [topic7 As Variant], [topic8 As Variant], [topic9 As Variant], [topic10 As Variant], [topic11 As Variant], [topic12 As Variant], [topic13 As Variant], [topic14 As Variant], [topic15 As Variant], [topic16 As Variant], [topic17 As Variant], [topic18 As Variant], [topic19 As Variant], [topic20 As Variant], [topic21 As Variant], [topic22 As Variant], [topic23 As Variant], [topic24 As Variant], [topic25 As Variant], [topic26 As Variant], [topic27 As Variant], [topic28 As Variant]) As Variant`  
  This method connects to a source to receive real-time data (RTD).
    - `progID As Variant` (required): A string representing the real-time server programmatic identifier.
    - `server As Variant` (required): A server name, Null string, or vbNullString constant.
    - `topic1 As Variant` (required): A String representing a topic.
- `Hex2Bin(Arg1 As Variant, [Arg2 As Variant]) As String`  
  Converts a hexadecimal number to binary.
    - `Arg1 As Variant` (required): Number - the hexadecimal number that you want to convert. Number cannot contain more than 10 characters. The most significant bit of number is the sign bit (40th bit from the right). The remaining 9 bits are magnitude bits. Negative numbers are represented using two's-complement notation.
    - `Arg2 As Variant` (optional): Places - the number of characters to use. If places is omitted, Hex2Bin uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).
- `Hex2Dec(Arg1 As Variant) As String`  
  Converts a hexadecimal number to decimal.
    - `Arg1 As Variant` (required): Number - the hexadecimal number that you want to convert. Number cannot contain more than 10 characters (40 bits). The most significant bit of number is the sign bit. The remaining 39 bits are magnitude bits. Negative numbers are represented using two's-complement notation.
- `Hex2Oct(Arg1 As Variant, [Arg2 As Variant]) As String`  
  Converts a hexadecimal number to octal.
    - `Arg1 As Variant` (required): Number - the hexadecimal number that you want to convert. Number cannot contain more than 10 characters. The most significant bit of number is the sign bit. The remaining 39 bits are magnitude bits. Negative numbers are represented using two's-complement notation.
    - `Arg2 As Variant` (optional): Places - the number of characters to use. If places is omitted, Hex2Oct uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).
- `Dec2Bin(Arg1 As Variant, [Arg2 As Variant]) As String`  
  Converts a decimal number to binary.
    - `Arg1 As Variant` (required): Number - the decimal integer that you want to convert. If number is negative, valid place values are ignored and Dec2Bin returns a 10-character (10-bit) binary number in which the most significant bit is the sign bit. The remaining 9 bits are magnitude bits. Negative numbers are represented using two's-complement notation.
    - `Arg2 As Variant` (optional): Places - the number of characters to use. If places is omitted, Dec2Bin uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).
- `Dec2Hex(Arg1 As Variant, [Arg2 As Variant]) As String`  
  Converts a decimal number to hexadecimal.
    - `Arg1 As Variant` (required): Number - the decimal integer that you want to convert. If number is negative, places is ignored and Dec2Hex returns a 10-character (40-bit) hexadecimal number in which the most significant bit is the sign bit. The remaining 39 bits are magnitude bits. Negative numbers are represented using two's-complement notation.
    - `Arg2 As Variant` (optional): Places - the number of characters to use. If places is omitted, Dec2Hex uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).
- `Dec2Oct(Arg1 As Variant, [Arg2 As Variant]) As String`  
  Converts a decimal number to octal.
    - `Arg1 As Variant` (required): Number - the decimal integer that you want to convert. If number is negative, places is ignored and Dec2Oct returns a 10-character (30-bit) octal number in which the most significant bit is the sign bit. The remaining 29 bits are magnitude bits. Negative numbers are represented using two's-complement notation.
    - `Arg2 As Variant` (optional): Places - the number of characters to use. If places is omitted, Dec2Oct uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).
- `Oct2Bin(Arg1 As Variant, [Arg2 As Variant]) As String`  
  Converts an octal number to binary.
    - `Arg1 As Variant` (required): Number - the octal number that you want to convert. Number may not contain more than 10 characters. The most significant bit of number is the sign bit. The remaining 29 bits are magnitude bits. Negative numbers are represented using two's-complement notation.
    - `Arg2 As Variant` (optional): Places - the number of characters to use. If places is omitted, Oct2Bin uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).
- `Oct2Hex(Arg1 As Variant, [Arg2 As Variant]) As String`  
  Converts an octal number to hexadecimal.
    - `Arg1 As Variant` (required): Number - the octal number that you want to convert. Number may not contain more than 10 octal characters (30 bits). The most significant bit of number is the sign bit. The remaining 29 bits are magnitude bits. Negative numbers are represented using two's-complement notation.
    - `Arg2 As Variant` (optional): Places - the number of characters to use. If places is omitted, Oct2Hex uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).
- `Oct2Dec(Arg1 As Variant) As String`  
  Converts an octal number to decimal.
    - `Arg1 As Variant` (required): Number - the octal number that you want to convert. Number may not contain more than 10 octal characters (30 bits). The most significant bit of number is the sign bit. The remaining 29 bits are magnitude bits. Negative numbers are represented using two's-complement notation.
- `Bin2Dec(Arg1 As Variant) As String`  
  Converts a binary number to decimal.
    - `Arg1 As Variant` (required): The binary number that you want to convert. Number cannot contain more than 10 characters (10 bits). The most significant bit of number is the sign bit. The remaining 9 bits are magnitude bits. Negative numbers are represented by using two's-complement notation.
- `Bin2Oct(Arg1 As Variant, [Arg2 As Variant]) As String`  
  Converts a binary number to octal.
    - `Arg1 As Variant` (required): The binary number that you want to convert. Number cannot contain more than 10 characters (10 bits). The most significant bit of number is the sign bit. The remaining 9 bits are magnitude bits. Negative numbers are represented using two's-complement notation.
    - `Arg2 As Variant` (optional): The number of characters to use. If places is omitted, Bin2Oct uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).
- `Bin2Hex(Arg1 As Variant, [Arg2 As Variant]) As String`  
  Converts a binary number to hexadecimal.
    - `Arg1 As Variant` (required): The binary number that you want to convert. Number cannot contain more than 10 characters (10 bits). The most significant bit of number is the sign bit. The remaining 9 bits are magnitude bits. Negative numbers are represented using two's-complement notation.
    - `Arg2 As Variant` (optional): The number of characters to use. If places is omitted, Bin2Hex uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).
- `ImSub(Arg1 As Variant, Arg2 As Variant) As String`  
  Returns the difference of two complex numbers in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber1 - the complex number from which to subtract inumber2.
    - `Arg2 As Variant` (required): Inumber2 - the complex number to subtract from inumber1.
- `ImDiv(Arg1 As Variant, Arg2 As Variant) As String`  
  Returns the quotient of two complex numbers in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber1 - the complex numerator or dividend.
    - `Arg2 As Variant` (required): Inumber2 - the complex denominator or divisor.
- `ImPower(Arg1 As Variant, Arg2 As Variant) As String`  
  Returns a complex number in x + yi or x + yj text format raised to a power.
    - `Arg1 As Variant` (required): Inumber - a complex number that you want to raise to a power.
    - `Arg2 As Variant` (required): Number - the power to which you want to raise the complex number.
- `ImAbs(Arg1 As Variant) As String`  
  Returns the absolute value (modulus) of a complex number in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber - a complex number for which you want the absolute value.
- `ImSqrt(Arg1 As Variant) As String`  
  Returns the square root of a complex number in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber - a complex number for which you want the square root.
- `ImLn(Arg1 As Variant) As String`  
  Returns the natural logarithm of a complex number in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber - a complex number for which you want the natural logarithm.
- `ImLog2(Arg1 As Variant) As String`  
  Returns the base-2 logarithm of a complex number in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber - a complex number for which you want the base-2 logarithm.
- `ImLog10(Arg1 As Variant) As String`  
  Returns the common logarithm (base 10) of a complex number in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber - a complex number for which you want the common logarithm.
- `ImSin(Arg1 As Variant) As String`  
  Returns the sine of a complex number in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber - a complex number for which you want the sine.
- `ImCos(Arg1 As Variant) As String`  
  Returns the cosine of a complex number in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber - a complex number for which you want the cosine.
- `ImExp(Arg1 As Variant) As String`  
  Returns the exponential of a complex number in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber - a complex number for which you want the exponential.
- `ImArgument(Arg1 As Variant) As String`  
  Returns the argument !Screenshot of the theta symbol. (theta), an angle expressed in radians, such that:
    - `Arg1 As Variant` (required): Inumber is a complex number for which you want the argument theta.
- `ImConjugate(Arg1 As Variant) As String`  
  Returns the complex conjugate of a complex number in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber - a complex number for which you want the conjugate.
- `Imaginary(Arg1 As Variant) As Double`  
  Returns the imaginary coefficient of a complex number in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber - a complex number for which you want the imaginary coefficient.
- `ImReal(Arg1 As Variant) As Double`  
  Returns the real coefficient of a complex number in x + yi or x + yj text format.
    - `Arg1 As Variant` (required): Inumber - a complex number for which you want the real coefficient.
- `Complex(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As String`  
  Converts real and imaginary coefficients into a complex number of the form x + yi or x + yj.
    - `Arg1 As Variant` (required): The real coefficient of the complex number.
    - `Arg2 As Variant` (required): The imaginary coefficient of the complex number.
    - `Arg3 As Variant` (optional): The suffix for the imaginary component of the complex number. If omitted, suffix is assumed to be "i".
- `ImSum(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As String`  
  Returns the sum of two or more complex numbers in x + yi or x + yj text format.
- `ImProduct(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As String`  
  Returns the product of 2 to 29 complex numbers in x + yi or x + yj text format.
- `SeriesSum(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Double`  
  Returns the sum of a power series based on the following formula:
    - `Arg1 As Variant` (required): X - the input value to the power series.
    - `Arg2 As Variant` (required): N - the initial power to which you want to raise x.
    - `Arg3 As Variant` (required): M - the step by which to increase n for each term in the series.
    - `Arg4 As Variant` (required): Coefficients - a set of coefficients by which each successive power of x is multiplied. The number of values in coefficients determines the number of terms in the power series. For example, if there are three values in coefficients, there will be three terms in the power series.
- `FactDouble(Arg1 As Variant) As Double`  
  Returns the double factorial of a number.
    - `Arg1 As Variant` (required): Number - the value for which to return the double factorial. If number is not an integer, it is truncated.
- `SqrtPi(Arg1 As Variant) As Double`  
  Returns the square root of (number * pi).
    - `Arg1 As Variant` (required): Number - the number by which pi is multiplied.
- `Quotient(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the integer portion of a division. Use this function when you want to discard the remainder of a division.
    - `Arg1 As Variant` (required): Numerator - the dividend.
    - `Arg2 As Variant` (required): Denominator - the divisor.
- `Delta(Arg1 As Variant, [Arg2 As Variant]) As Double`  
  Tests whether two values are equal. Returns 1 if number1 = number2; otherwise, returns 0.
    - `Arg1 As Variant` (required): Number1 - the first number.
    - `Arg2 As Variant` (optional): Number2 - the second number. If omitted, number2 is assumed to be zero.
- `GeStep(Arg1 As Variant, [Arg2 As Variant]) As Double`  
  Returns 1 if number  step; otherwise, returns 0 (zero). Use this function to filter a set of values. For example, by summing several GeStep functions, you calculate the count of values that exceed a threshold.
    - `Arg1 As Variant` (required): Number - the value to test against step.
    - `Arg2 As Variant` (optional): Step - the threshold value. If you omit a value for step, GeStep uses zero.
- `IsEven(Arg1 As Variant) As Boolean`  
  Checks the type of value and returns True or False depending on whether the value is even.
    - `Arg1 As Variant` (required): Value - the value that you want tested. Value can be a blank (empty cell), error, logical, text, number, or reference value, or a name referring to any of these, that you want to test.
- `IsOdd(Arg1 As Variant) As Boolean`  
  Checks the type of value and returns True or False depending on whether the value is odd.
    - `Arg1 As Variant` (required): Value - the value that you want tested. Value can be a blank (empty cell), error, logical, text, number, or reference value, or a name referring to any of these, that you want to test.
- `MRound(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns a number rounded to the desired multiple.
    - `Arg1 As Variant` (required): Number - the value to round.
    - `Arg2 As Variant` (required): Multiple - the multiple to which you want to round number.
- `Erf(Arg1 As Variant, [Arg2 As Variant]) As Double`  
  Returns the error function integrated between lower_limit and upper_limit.
    - `Arg1 As Variant` (required): Lower_limit - the lower bound for integrating Erf.
    - `Arg2 As Variant` (optional): Upper_limit - the upper bound for integrating Erf. If omitted, Erf integrates between zero and lower_limit.
- `ErfC(Arg1 As Variant) As Double`  
  Returns the complementary Erf function integrated between the specified parameter and infinity.
    - `Arg1 As Variant` (required): The first argument.
- `BesselJ(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the Bessel function.
    - `Arg1 As Variant` (required): The value at which to evaluate the function.
    - `Arg2 As Variant` (required): The order of the Bessel function. If n is not an integer, it is truncated.
- `BesselK(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the modified Bessel function, which is equivalent to the Bessel functions evaluated for purely imaginary arguments.
    - `Arg1 As Variant` (required): The value at which to evaluate the function.
    - `Arg2 As Variant` (required): The order of the function. If n is not an integer, it is truncated.
- `BesselY(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the Bessel function, which is also called the Weber function or the Neumann function.
    - `Arg1 As Variant` (required): The value at which to evaluate the function.
    - `Arg2 As Variant` (required): The order of the function. If n is not an integer, it is truncated.
- `BesselI(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the modified Bessel function, which is equivalent to the Bessel function evaluated for purely imaginary arguments.
    - `Arg1 As Variant` (required): The value at which to evaluate the function.
    - `Arg2 As Variant` (required): The order of the Bessel function. If n is not an integer, it is truncated.
- `Xirr(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As Double`  
  Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic. To calculate the internal rate of return for a series of periodic cash flows, use the Irr function.
    - `Arg1 As Variant` (required): Values - a series of cash flows that corresponds to a schedule of payments in dates. The first payment is optional and corresponds to a cost or payment that occurs at the beginning of the investment. If the first value is a cost or payment, it must be a negative value. All succeeding payments are discounted based on a 365-day year. The series of values must contain at least one positive and one negative value.
    - `Arg2 As Variant` (required): Dates - a schedule of payment dates that corresponds to the cash flow payments. The first payment date indicates the beginning of the schedule of payments. All other dates must be later than this date, but they may occur in any order. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
    - `Arg3 As Variant` (optional): Guess - a number that you guess is close to the result of Xirr.
- `Xnpv(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the net present value for a schedule of cash flows that is not necessarily periodic. Read/write Double.
    - `Arg1 As Variant` (required): A series of cash flows that corresponds to a schedule of payments in dates. The first payment is optional and corresponds to a cost or payment that occurs at the beginning of the investment.
    - `Arg2 As Variant` (required): A schedule of payment dates that corresponds to the cash flow payments. The first payment date indicates the beginning of the schedule of payments. All other dates must be later than this date, but they may occur in any order.
- `PriceMat(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, [Arg6 As Variant]) As Double`  
  Returns the price per $100 face value of a security that pays interest at maturity.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Issue - the security's issue date, expressed as a serial date number.
    - `Arg4 As Variant` (required): Rate - the security's interest rate at date of issue.
    - `Arg5 As Variant` (required): Yld - the security's annual yield.
    - `Arg6 As Variant` (optional): Basis - the type of day count basis to use.
- `YieldMat(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, [Arg6 As Variant]) As Double`  
  Returns the annual yield of a security that pays interest at maturity.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Issue - the security's issue date, expressed as a serial date number.
    - `Arg4 As Variant` (required): Rate - the security's interest rate at date of issue.
    - `Arg5 As Variant` (required): Pr - the security's price per $100 face value.
    - `Arg6 As Variant` (optional): Basis - the type of day count basis to use.
- `IntRate(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, [Arg5 As Variant]) As Double`  
  Returns the interest rate for a fully invested security.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Investment - the amount invested in the security.
    - `Arg4 As Variant` (required): Redemption - the amount to be received at maturity.
    - `Arg5 As Variant` (optional): Basis - the type of day count basis to use.
- `Received(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, [Arg5 As Variant]) As Double`  
  Returns the amount received at maturity for a fully invested security.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Investment - the amount invested in the security.
    - `Arg4 As Variant` (required): Discount - the security's discount rate.
    - `Arg5 As Variant` (optional): Basis - the type of day count basis to use.
- `Disc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, [Arg5 As Variant]) As Double`  
  Returns the discount rate for a security.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Pr - the security's price per $100 face value.
    - `Arg4 As Variant` (required): Redemption - the security's redemption value per $100 face value.
    - `Arg5 As Variant` (optional): Basis - the type of day count basis to use.
- `PriceDisc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, [Arg5 As Variant]) As Double`  
  Returns the price per $100 face value of a discounted security.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Discount - the security's discount rate.
    - `Arg4 As Variant` (required): Redemption - the security's redemption value per $100 face value.
    - `Arg5 As Variant` (optional): Basis - the type of day count basis to use.
- `YieldDisc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, [Arg5 As Variant]) As Double`  
  Returns the annual yield for a discounted security.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Pr - the security's price per $100 face value.
    - `Arg4 As Variant` (required): Redemption - the security's redemption value per $100 face value.
    - `Arg5 As Variant` (optional): Basis - the type of day count basis to use.
- `TBillEq(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As Double`  
  Returns the bond-equivalent yield for a Treasury bill.
    - `Arg1 As Variant` (required): Settlement - the Treasury bill's settlement date. The security settlement date is the date after the issue date when the Treasury bill is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the Treasury bill's maturity date. The maturity date is the date when the Treasury bill expires.
    - `Arg3 As Variant` (optional): Discount - the Treasury bill's discount rate.
- `TBillPrice(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As Double`  
  Returns the price per $100 face value for a Treasury bill.
    - `Arg1 As Variant` (required): Settlement - the Treasury bill's settlement date. The security settlement date is the date after the issue date when the Treasury bill is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the Treasury bill's maturity date. The maturity date is the date when the Treasury bill expires.
    - `Arg3 As Variant` (optional): Discount - the Treasury bill's discount rate.
- `TBillYield(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As Double`  
  Returns the yield for a Treasury bill.
    - `Arg1 As Variant` (required): Settlement - the Treasury bill's settlement date. The security settlement date is the date after the issue date when the Treasury bill is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the Treasury bill's maturity date. The maturity date is the date when the Treasury bill expires.
    - `Arg3 As Variant` (optional): Pr - the Treasury bill's price per $100 face value.
- `Price(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, [Arg7 As Variant]) As Double`  
  Returns the price per $100 face value of a security that pays periodic interest.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Rate - the security's annual coupon rate.
    - `Arg4 As Variant` (required): Yld - the security's annual yield.
    - `Arg5 As Variant` (required): Redemption - the security's redemption value per $100 face value.
    - `Arg6 As Variant` (required): Frequency - the number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg7 As Variant` (optional): Basis - the type of day count basis to use.
- `DollarDe(Arg1 As Variant, Arg2 As Variant) As Double`  
  Converts a dollar price expressed as a fraction into a dollar price expressed as a decimal number. Use DollarDe to convert fractional dollar numbers, such as securities prices, to decimal numbers.
    - `Arg1 As Variant` (required): Fractional_dollar - a number expressed as a fraction.
    - `Arg2 As Variant` (required): Fraction - the integer to use in the denominator of the fraction.
- `DollarFr(Arg1 As Variant, Arg2 As Variant) As Double`  
  Converts a dollar price expressed as a decimal number into a dollar price expressed as a fraction. Use DollarFr to convert decimal numbers to fractional dollar numbers, such as securities prices.
    - `Arg1 As Variant` (required): Decimal_dollar - a decimal number.
    - `Arg2 As Variant` (required): Fraction - the integer to use in the denominator of a fraction.
- `Nominal(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the nominal annual interest rate, given the effective rate and the number of compounding periods per year.
    - `Arg1 As Variant` (required): Effect_rate - the effective interest rate.
    - `Arg2 As Variant` (required): Npery - the number of compounding periods per year.
- `Effect(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the effective annual interest rate, given the nominal annual interest rate and the number of compounding periods per year.
    - `Arg1 As Variant` (required): Nominal_rate - the nominal interest rate.
    - `Arg2 As Variant` (required): Npery - the number of compounding periods per year.
- `CumPrinc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant) As Double`  
  Returns the cumulative principal paid on a loan between start_period and end_period.
    - `Arg1 As Variant` (required): The interest rate.
    - `Arg2 As Variant` (required): The total number of payment periods.
    - `Arg3 As Variant` (required): The present value.
    - `Arg4 As Variant` (required): The first period in the calculation. Payment periods are numbered beginning with 1.
    - `Arg5 As Variant` (required): The last period in the calculation.
    - `Arg6 As Variant` (required): The timing of the payment.
- `CumIPmt(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant) As Double`  
  Returns the cumulative interest paid on a loan between start_period and end_period.
    - `Arg1 As Variant` (required): The interest rate.
    - `Arg2 As Variant` (required): The total number of payment periods.
    - `Arg3 As Variant` (required): The present value.
    - `Arg4 As Variant` (required): The first period in the calculation. Payment periods are numbered beginning with 1.
    - `Arg5 As Variant` (required): The last period in the calculation.
    - `Arg6 As Variant` (required): The timing of the payment.
- `EDate(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the serial number that represents the date that is the indicated number of months before or after a specified date (the start_date). Use EDate to calculate maturity dates or due dates that fall on the same day of the month as the date of issue.
    - `Arg1 As Variant` (required): Start_date - a date that represents the start date. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
    - `Arg2 As Variant` (required): Months - the number of months before or after start_date. A positive value for months yields a future date; a negative value yields a past date.
- `EoMonth(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the serial number for the last day of the month that is the indicated number of months before or after start_date. Use EoMonth to calculate maturity dates or due dates that fall on the last day of the month.
    - `Arg1 As Variant` (required): Start_date - a date that represents the starting date. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
    - `Arg2 As Variant` (required): Months - the number of months before or after start_date. A positive value for months yields a future date; a negative value yields a past date.
- `YearFrac(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As Double`  
  Calculates the fraction of the year represented by the number of whole days between two dates (the start_date and the end_date). Use the YearFrac worksheet function to identify the proportion of a whole year's benefits or obligations to assign to a specific term.
    - `Arg1 As Variant` (required): Start_date - a date that represents the start date.
    - `Arg2 As Variant` (required): End_date - a date that represents the end date.
    - `Arg3 As Variant` (optional): Basis - the type of day count basis to use.
- `CoupDayBs(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant]) As Double`  
  Returns the number of days from the beginning of the coupon period to the settlement date.
    - `Arg1 As Variant` (required): The security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): The security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): The number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg4 As Variant` (optional): The type of day count basis to use.
- `CoupDays(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant]) As Double`  
  Returns the number of days in the coupon period that contain the settlement date.
    - `Arg1 As Variant` (required): The security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): The security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): The number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg4 As Variant` (optional): The type of day count basis to use.
- `CoupDaysNc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant]) As Double`  
  Returns the number of days from the settlement date to the next coupon date.
    - `Arg1 As Variant` (required): The security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): The security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): The number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg4 As Variant` (optional): The type of day count basis to use.
- `CoupNcd(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant]) As Double`  
  Returns a number that represents the next coupon date after the settlement date.
    - `Arg1 As Variant` (required): The security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): The security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): The number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg4 As Variant` (optional): The type of day count basis to use.
- `CoupNum(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant]) As Double`  
  Returns the number of coupons payable between the settlement date and maturity date, rounded up to the nearest whole coupon.
    - `Arg1 As Variant` (required): The security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): The security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): The number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg4 As Variant` (optional): The type of day count basis to use.
- `CoupPcd(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant]) As Double`
    - `Arg1 As Variant` (required): The security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): The security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): The number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg4 As Variant` (optional): The type of day count basis to use.
- `Duration(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, [Arg6 As Variant]) As Double`  
  Returns the Macauley duration for an assumed par value of $100. Duration is defined as the weighted average of the present value of the cash flows and is used as a measure of a bond price's response to changes in yield.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Coupon - the security's annual coupon rate.
    - `Arg4 As Variant` (required): Yld - the security's annual yield.
    - `Arg5 As Variant` (required): Frequency - the number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg6 As Variant` (optional): Basis - the type of day count basis to use.
- `MDuration(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, [Arg6 As Variant]) As Double`  
  Returns the modified Macauley duration for a security with an assumed par value of $100.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Coupon - the security's annual coupon rate.
    - `Arg4 As Variant` (required): Yld - the security's annual yield.
    - `Arg5 As Variant` (required): Frequency - the number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg6 As Variant` (optional): Basis - the type of day count basis to use
- `OddLPrice(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Arg7 As Variant, [Arg8 As Variant]) As Double`  
  Returns the price per $100 face value of a security having an odd (short or long) last coupon period.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Last_interest - the security's last coupon date.
    - `Arg4 As Variant` (required): Rate - the security's interest rate.
    - `Arg5 As Variant` (required): Yld - the security's annual yield.
    - `Arg6 As Variant` (required): Redemption - the security's redemption value per $100 face value.
    - `Arg7 As Variant` (required): Frequency - the number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg8 As Variant` (optional): Basis - the type of day count basis to use.
- `OddLYield(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Arg7 As Variant, [Arg8 As Variant]) As Double`  
  Returns the yield of a security that has an odd (short or long) last period.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Last_interest - the security's last coupon date.
    - `Arg4 As Variant` (required): Rate - the security's interest rate.
    - `Arg5 As Variant` (required): Pr - the security's price.
    - `Arg6 As Variant` (required): Redemption - the security's redemption value per $100 face value.
    - `Arg7 As Variant` (required): Frequency - the number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg8 As Variant` (optional): Basis - the type of day count basis to use.
- `OddFPrice(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Arg7 As Variant, Arg8 As Variant, [Arg9 As Variant]) As Double`  
  Returns the price per $100 face value of a security having an odd (short or long) first period.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Issue - the security's issue date.
    - `Arg4 As Variant` (required): First_coupon - the security's first coupon date.
    - `Arg5 As Variant` (required): Rate - the security's interest rate.
    - `Arg6 As Variant` (required): Yld - the security's annual yield.
    - `Arg7 As Variant` (required): Redemption - the security's redemption value per $100 face value.
    - `Arg8 As Variant` (required): Frequency - the number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg9 As Variant` (optional): Basis - the type of day count basis to use.
- `OddFYield(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Arg7 As Variant, Arg8 As Variant, [Arg9 As Variant]) As Double`  
  Returns the yield of a security that has an odd (short or long) first period.
    - `Arg1 As Variant` (required): Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.
    - `Arg2 As Variant` (required): Maturity - the security's maturity date. The maturity date is the date when the security expires.
    - `Arg3 As Variant` (required): Issue - the security's issue date.
    - `Arg4 As Variant` (required): First_coupon - the security's first coupon date.
    - `Arg5 As Variant` (required): Rate - the security's interest rate.
    - `Arg6 As Variant` (required): Pr - the security's price.
    - `Arg7 As Variant` (required): Redemption - the security's redemption value per $100 face value.
    - `Arg8 As Variant` (required): Frequency - the number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.
    - `Arg9 As Variant` (optional): Basis - the type of day count basis to use.
- `RandBetween(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns a random integer number between the numbers that you specify. A new random integer number is returned every time the worksheet is calculated.
    - `Arg1 As Variant` (required): Bottom - the smallest integer RandBetween will return.
    - `Arg2 As Variant` (required): Top - the largest integer RandBetween will return.
- `WeekNum(Arg1 As Variant, [Arg2 As Variant]) As Double`  
  Returns a number that indicates where the week falls numerically within a year.
    - `Arg1 As Variant` (required): Serial_num - a date within the week. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
    - `Arg2 As Variant` (optional): Return_type - a number that determines on which day the week begins. The default is 1.
- `AmorDegrc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, [Arg7 As Variant]) As Double`  
  Returns the depreciation for each accounting period. This function is provided for the French accounting system.
    - `Arg1 As Variant` (required): The cost of the asset.
    - `Arg2 As Variant` (required): The date of the purchase of the asset.
    - `Arg3 As Variant` (required): The date of the end of the first period.
    - `Arg4 As Variant` (required): The salvage value at the end of the life of the asset.
    - `Arg5 As Variant` (required): The period.
    - `Arg6 As Variant` (required): The rate of depreciation.
    - `Arg7 As Variant` (optional): The year basis to be used.
- `AmorLinc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, [Arg7 As Variant]) As Double`  
  Returns the depreciation for each accounting period. This function is provided for the French accounting system.
    - `Arg1 As Variant` (required): The cost of the asset.
    - `Arg2 As Variant` (required): The date of the purchase of the asset.
    - `Arg3 As Variant` (required): The date of the end of the first period.
    - `Arg4 As Variant` (required): The salvage value at the end of the life of the asset.
    - `Arg5 As Variant` (required): The period.
    - `Arg6 As Variant` (required): The rate of depreciation.
    - `Arg7 As Variant` (optional): The year basis to be used.
- `Convert(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Double`  
  Converts a number from one measurement system to another. For example, Convert can translate a table of distances in miles to a table of distances in kilometers.
    - `Arg1 As Variant` (required): The value in from_units to convert.
    - `Arg2 As Variant` (required): The units for number.
    - `Arg3 As Variant` (required): The units for the result. Convert accepts the following text values (in quotation marks) for from_unit and to_unit, which are listed in the Remarks section.
- `AccrInt(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, [Arg7 As Variant]) As Double`  
  Returns the accrued interest for a security that pays periodic interest.
    - `Arg1 As Variant` (required): Issue date - Security's issue date.
    - `Arg2 As Variant` (required): First interest - Security's first interest date.
    - `Arg3 As Variant` (required): Settlement - Security's settlement date
    - `Arg4 As Variant` (required): Rate - Security's annual coupon rate.
    - `Arg5 As Variant` (required): Par - Security's par value.
    - `Arg6 As Variant` (required): Frequency - Number of coupon payments per year.
    - `Arg7 As Variant` (optional): Basis - The type of day count basis to use.
- `AccrIntM(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, [Arg5 As Variant]) As Double`  
  Returns the accrued interest for a security that pays interest at maturity.
    - `Arg1 As Variant` (required): The security's issue date.
    - `Arg2 As Variant` (required): The security's maturity date.
    - `Arg3 As Variant` (required): The security's annual coupon rate.
    - `Arg4 As Variant` (required): The security's par value. If you omit par, ACCRINTM uses $1,000.
    - `Arg5 As Variant` (optional): The type of day count basis to use.
- `WorkDay(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As Double`  
  Returns a number that represents a date that is the indicated number of working days before or after a date (the starting date). Working days exclude weekends and any dates identified as holidays. Use WorkDay to exclude weekends or holidays when you calculate invoice due dates, expected delivery times, or the number of days of work performed.
    - `Arg1 As Variant` (required): Start_date - a date that represents the start date.
    - `Arg2 As Variant` (required): Days - the number of nonweekend and nonholiday days before or after start_date. A positive value for days yields a future date; a negative value yields a past date.
    - `Arg3 As Variant` (optional): Holidays - an optional list of one or more dates to exclude from the working calendar, such as state and federal holidays and floating holidays. The list can be either a range of cells that contain the dates or an array constant of the serial numbers that represent the dates.
- `NetworkDays(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As Double`  
  Returns the number of whole working days between start_date and end_date. Working days exclude weekends and any dates identified in holidays. Use NetworkDays to calculate employee benefits that accrue based on the number of days worked during a specific term.
    - `Arg1 As Variant` (required): Start_date - a date that represents the start date.
    - `Arg2 As Variant` (required): End_date - a date that represents the end date.
    - `Arg3 As Variant` (optional): Holidays - an optional range of one or more dates to exclude from the working calendar, such as state and federal holidays and floating holidays. The list can be either a range of cells that contains the dates or an array constant of the serial numbers that represent the dates.
- `Gcd(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the greatest common divisor of two or more integers. The greatest common divisor is the largest integer that divides both number1 and number2 without a remainder.
- `MultiNomial(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the ratio of the factorial of a sum of values to the product of factorials.
- `Lcm(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the least common multiple of integers. The least common multiple is the smallest positive integer that is a multiple of all integer arguments number1, number2, and so on. Use Lcm to add fractions with different denominators.
    - `Arg1 As Variant` (required): Number1, number2... - 1 to 29 values for which you want the least common multiple. If value is not an integer, it is truncated.
- `FVSchedule(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the future value of an initial principal after applying a series of compound interest rates. Use FVSchedule to calculate the future value of an investment with a variable or adjustable rate.
    - `Arg1 As Variant` (required): Principal - the present value.
    - `Arg2 As Variant` (required): Schedule - an array of interest rates to apply.
- `SumIfs(Arg1 As Range, Arg2 As Range, Arg3 As Variant, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant]) As Double`  
  Adds the cells in a range that meet multiple criteria.
    - `Arg1 As Range` (required): Sum_range - the range to sum.
    - `Arg2 As Range` (required): Criteria_range1, criteria_range2... - One or more ranges in which to evaluate the associated criteria.
- `CountIfs(Arg1 As Range, Arg2 As Variant, [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Counts the number of cells within a range that meet multiple criteria.
    - `Arg1 As Range` (required): One or more ranges in which to evaluate the associated criteria.
- `AverageIf(Arg1 As Range, Arg2 As Variant, [Arg3 As Variant]) As Double`  
  Returns the average (arithmetic mean) of all the cells in a range that meet a given criteria.
    - `Arg1 As Range` (required): One or more cells to average.
    - `Arg2 As Variant` (required): The criteria in the form of a number, expression, cell reference, or text that defines which cells are averaged. For example, criteria can be expressed as 32, "32", ">32", "apples", or B4.
    - `Arg3 As Variant` (optional): The actual set of cells to average. If omitted, range is used.
- `AverageIfs(Arg1 As Range, Arg2 As Range, Arg3 As Variant, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant]) As Double`  
  Returns the average (arithmetic mean) of all cells that meet multiple criteria.
- `IfError(Arg1 As Variant, Arg2 As Variant) As Variant`  
  Returns a value that you specify if a formula evaluates to an error; otherwise, returns the result of the formula. Use the IfError function to trap and handle errors in a formula.
    - `Arg1 As Variant` (required): Value - the argument that is checked for an error.
    - `Arg2 As Variant` (required): Value_if_error - the value to return if the formula evaluates to an error. The following error types are evaluated: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!.
- `Aggregate(Arg1 As Double, Arg2 As Double, Arg3 As Range, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns an aggregate in a list or database.
    - `Arg1 As Double` (required): Function_num - A number from 1 to 19 that specifies which function to use.<table><tr><th>Function_num</th><th>Function</th></tr><tr><td>1</td><td>AVERAGE</td></tr><tr><td>2</td><td>COUNT</td></tr><tr><td>3</td><td>COUNTA</td></tr><tr><td>4</td><td>MAX</td></tr><tr><td>5</td><td>MIN</td></tr><tr><td>6</td><td>PRODUCT</td></tr><tr><td>7</td><td>STDEV.S</td></tr><tr><td>8</td><td>STDEV.P</td></tr><tr><td>9</td><td>SUM</td></tr><tr><td>10</td><td>VAR.S</td></tr><tr><td>11</td><td>VAR.P</td></tr><tr><td>12</td><td>MEDIAN</td></tr><tr><td>13</td><td>MODE.SNGL</td></tr><tr><td>14</td><td>LARGE</td></tr><tr><td>15</td><td>SMALL</td></tr><tr><td>16</td><td>PERCENTILE.INC </td></tr><tr><td>17</td><td>QUARTILE.INC</td></tr><tr><td>18</td><td>PERCENTILE.EXC</td></tr><tr><td>19</td><td>QUARTILE.EXC</td></tr></table>
    - `Arg2 As Double` (required): Option - A numerical value that determines which values to ignore in the evaluation range for the function.<table><tr><th>Option</th><th>Behavior</th> </tr><tr><td>0 or omitted</td><td>Ignore nested SUBTOTAL and AGGREGATE functions</td> </tr><tr><td>1</td><td>Ignore hidden rows, nested SUBTOTAL and AGGREGATE functions</td> </tr><tr><td>2</td><td>Ignore error values, nested SUBTOTAL and AGGREGATE functions</td> </tr><tr><td>3</td><td>Ignore hidden rows, error values, nested SUBTOTAL and AGGREGATE functions</td> </tr><tr><td>4</td><td>Ignore nothing</td> </tr><tr><td>5</td><td>Ignore hidden rows</td> </tr><tr><td>6</td><td>Ignore error values</td> </tr><tr><td>7</td><td>Ignore hidden rows and error values</td> </tr></table>
    - `Arg3 As Range` (required): Ref1 - The first numeric argument for functions that take multiple numeric arguments for which you want the aggregate value.
- `Confidence_Norm(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns a value that you can use to construct a confidence interval for a population mean.
    - `Arg1 As Double` (required): The significance level used to compute the confidence level. The confidence level equals 100*(1 - alpha)%, or in other words, an alpha of 0.05 indicates a 95 percent confidence level.
    - `Arg2 As Double` (required): The population standard deviation for the data range; is assumed to be known.
    - `Arg3 As Double` (required): The sample size.
- `Confidence_T(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the confidence interval for a population mean, using a Student's t distribution.
    - `Arg1 As Double` (required): Alpha - The significance level used to compute the confidence level. The confidence level equals 100*(1 - alpha)%, or in other words, an alpha of 0.05 indicates a 95 percent confidence level.
    - `Arg2 As Double` (required): Standard_dev - The population standard deviation for the data range; is assumed to be known.
    - `Arg3 As Double` (required): Size - The sample size.
- `ChiSq_Test(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the test for independence.
    - `Arg1 As Variant` (required): The range of data that contains observations to test against expected values.
    - `Arg2 As Variant` (required): The range of data that contains the ratio of the product of row totals and column totals to the grand total.
- `F_Test(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the result of an F-test. An F-test returns the two-tailed probability that the variances in array1 and array2 are not significantly different. Use this function to determine whether two samples have different variances. For example, given test scores from public and private schools, you can test whether these schools have different levels of test score diversity.
    - `Arg1 As Variant` (required): Array1 - the first array or range of data.
    - `Arg2 As Variant` (required): Array2 - the second array or range of data.
- `Covariance_P(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns population covariance, the average of the products of deviations for each data point pair.
    - `Arg1 As Variant` (required): The first cell range of integers.
    - `Arg2 As Variant` (required): The second cell range of integers.
- `Covariance_S(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the sample covariance, the average of the products of deviations for each data point pair in two data sets.
    - `Arg1 As Variant` (required): Array1 - The first cell range of integers.
    - `Arg2 As Variant` (required): Array2 - The second cell range of integers.
- `Expon_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Boolean) As Double`  
  Returns the exponential distribution. Use Expon_Dist to model the time between events, such as how long an automated bank teller takes to deliver cash. For example, you can use Expon_Dist to determine the probability that the process takes at most 1 minute.
    - `Arg1 As Double` (required): X - the value of the function.
    - `Arg2 As Double` (required): Lambda - the parameter value.
    - `Arg3 As Boolean` (required): Cumulative - a logical value that indicates which form of the exponential function to provide. If cumulative is True, EXPONDIST returns the cumulative distribution function; if False, it returns the probability density function.
- `Gamma_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean) As Double`  
  Returns the gamma distribution. Use this function to study variables that may have a skewed distribution. The gamma distribution is commonly used in queuing analysis.
    - `Arg1 As Double` (required): X - the value at which you want to evaluate the distribution.
    - `Arg2 As Double` (required): Alpha - a parameter to the distribution.
    - `Arg3 As Double` (required): Beta - a parameter to the distribution. If beta = 1, Gamma_Dist returns the standard gamma distribution.
    - `Arg4 As Boolean` (required): Cumulative - a logical value that determines the form of the function. If cumulative is True, Gamma_Dist returns the cumulative distribution function; if False, it returns the probability density function.
- `Gamma_Inv(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the inverse of the gamma cumulative distribution. If p = GAMMA_DIST(x,...), then GAMMA_INV(p,...) = x.
    - `Arg1 As Double` (required): Probability - the probability associated with the gamma distribution.
    - `Arg2 As Double` (required): Alpha - a parameter to the distribution.
    - `Arg3 As Double` (required): Beta - a parameter to the distribution. If beta = 1, Gamma_Inv returns the standard gamma distribution.
- `Mode_Mult(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Variant`  
  Returns a vertical array of the most frequently occurring, or repetitive, values in an array or range of data.
    - `Arg1 As Variant` (required): Number1 - The first number argument for which you want to calculate the mode.
- `Mode_Sngl(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the most frequently occurring, or repetitive, value in an array or range of data.
- `Norm_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean) As Double`  
  Returns the normal distribution for the specified mean and standard deviation. This function has a wide range of applications in statistics, including hypothesis testing.
    - `Arg1 As Double` (required): X - The value for which you want the distribution.
    - `Arg2 As Double` (required): Mean - The arithmetic mean of the distribution.
    - `Arg3 As Double` (required): Standard_dev - The standard deviation of the distribution.
    - `Arg4 As Boolean` (required): Cumulative - A logical value that determines the form of the function. If cumulative is True, Norm_Dist returns the cumulative distribution function; if False, it returns the probability mass function.
- `Norm_Inv(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.
    - `Arg1 As Double` (required): Probability - A probability corresponding to the normal distribution.
    - `Arg2 As Double` (required): Mean - The arithmetic mean of the distribution.
    - `Arg3 As Double` (required): Standard_dev - The standard deviation of the distribution.
- `Percentile_Exc(Arg1 As Variant, Arg2 As Double) As Double`  
  Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.
    - `Arg1 As Variant` (required): Array - The array or range of data that defines relative standing.
    - `Arg2 As Double` (required): K - The percentile value in the range 0..1, exclusive.
- `Percentile_Inc(Arg1 As Variant, Arg2 As Double) As Double`  
  Returns the k-th percentile of values in a range. Use this function to establish a threshold of acceptance. For example, you can examine candidates who score above the 90th percentile.
    - `Arg1 As Variant` (required): Array - The array or range of data that defines relative standing.
    - `Arg2 As Double` (required): K - The percentile value in the range 0..1, inclusive.
- `PercentRank_Exc(Arg1 As Variant, Arg2 As Double, [Arg3 As Variant]) As Double`  
  Returns the rank of a value in a data set as a percentage (0..1, exclusive) of the data set.
    - `Arg1 As Variant` (required): Array - The array or range of data with numeric values that defines relative standing.
    - `Arg2 As Double` (required): X - The value for which you want to know the rank.
    - `Arg3 As Variant` (optional): Significance - A value that identifies the number of significant digits for the returned percentage value. If omitted, PercentRank_Exc uses three digits (0.xxx).
- `PercentRank_Inc(Arg1 As Variant, Arg2 As Double, [Arg3 As Variant]) As Double`  
  Returns the rank of a value in a data set as a percentage (0..1, inclusive) of the data set. This function can be used to evaluate the relative standing of a value within a data set. For example, you can use PercentRank_Inc to evaluate the standing of an aptitude test score among all scores for the test.
    - `Arg1 As Variant` (required): Array - The array or range of data with numeric values that defines relative standing.
    - `Arg2 As Double` (required): X - The value for which you want to know the rank.
    - `Arg3 As Variant` (optional): Significance - An optional value that identifies the number of significant digits for the returned percentage value. If omitted, PercentRank_Inc uses three digits (0.xxx).
- `Poisson_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Boolean) As Double`  
  Returns the Poisson distribution. A common application of the Poisson distribution is predicting the number of events over a specific time, such as the number of cars arriving at a toll plaza in one minute.
    - `Arg1 As Double` (required): X - The number of events.
    - `Arg2 As Double` (required): Mean - The expected numeric value.
    - `Arg3 As Boolean` (required): Cumulative - A logical value that determines the form of the probability distribution returned. If cumulative is True, Poisson_Dist returns the cumulative Poisson probability that the number of random events occurring will be between 0 (zero) and x inclusive; if False, it returns the Poisson probability mass function that the number of events occurring will be exactly x.
- `Quartile_Exc(Arg1 As Variant, Arg2 As Double) As Double`  
  Returns the quartile of the data set, based on percentile values from 0..1, exclusive.
    - `Arg1 As Variant` (required): Array - The array or cell range of numeric values for which you want the quartile value.
    - `Arg2 As Double` (required): Quart - The value to return.
- `Quartile_Inc(Arg1 As Variant, Arg2 As Double) As Double`  
  Returns the quartile of a data set based on percentile values from 0..1, inclusive. Quartiles often are used in sales and survey data to divide populations into groups. For example, you can use Quartile_Inc to find the top 25 percent of incomes in a population.
    - `Arg1 As Variant` (required): Array - the array or cell range of numeric values for which you want the quartile value.
    - `Arg2 As Double` (required): Quart - The value to return.
- `Rank_Avg(Arg1 As Double, Arg2 As Range, [Arg3 As Variant]) As Double`  
  Returns the rank of a number in a list of numbers; that is, its size relative to other values in the list. If more than one value has the same rank, the average rank is returned.
    - `Arg1 As Double` (required): Number - The number whose rank you want to find.
    - `Arg2 As Range` (required): Ref - An array of, or a reference to, a list of numbers. Non-numeric values in reference are ignored.
    - `Arg3 As Variant` (optional): Order - A number that specifies how to rank number. If the order is 0 (zero) or omitted, Microsoft Excel ranks the number as if the reference were a list sorted in descending order. If the order is any non-zero value, Excel ranks number as if the reference were a list sorted in ascending order.
- `Rank_Eq(Arg1 As Double, Arg2 As Range, [Arg3 As Variant]) As Double`  
  Returns the rank of a number in a list of numbers. The rank of a number is its size relative to other values in a list. If you were to sort the list, the rank of the number would be its position.
    - `Arg1 As Double` (required): Number - The number whose rank you want to find.
    - `Arg2 As Range` (required): Ref - An array of, or a reference to, a list of numbers. Non-numeric values in reference are ignored.
    - `Arg3 As Variant` (optional): Order - A number that specifies how to rank the number.
- `StDev_S(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Estimates standard deviation based on a sample. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).
- `StDev_P(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Calculates standard deviation based on the entire population given as arguments. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).
- `T_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Boolean) As Double`  
  Returns a Student t-distribution where a numeric value (x) is a calculated value of t for which the Percentage Points are computed.
    - `Arg1 As Double` (required): X - The numeric value at which to evaluate the distribution.
    - `Arg2 As Double` (required): Deg_freedom - An integer that indicates the number of degrees of freedom.
    - `Arg3 As Boolean` (required): Cumulative - A logical value that determines the form of the function. If cumulative is True, T_Dist returns the cumulative distribution function; if False, it returns the probability density function.
- `T_Dist_2T(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the two-tailed Student t-distribution.
    - `Arg1 As Double` (required): X - The numeric value at which to evaluate the distribution.
    - `Arg2 As Double` (required): Deg_freedom - An integer that indicates the number of degrees of freedom.
- `T_Dist_RT(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the right-tailed Student t-distribution where a numeric value (x) is a calculated value of t for which the Percentage Points are to be computed. The t-distribution is used in the hypothesis testing of small sample data sets. Use this function in place of a table of critical values for the t-distribution.
    - `Arg1 As Double` (required): X - The numeric value at which to evaluate the distribution.
    - `Arg2 As Double` (required): Degrees_freedom - An integer that indicates the number of degrees of freedom.
- `T_Inv(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the left-tailed inverse of the Student t-distribution.
    - `Arg1 As Double` (required): Probability - The probability associated with the Student t-distribution.
    - `Arg2 As Double` (required): Deg_freedom - The number of degrees of freedom with which to characterize the distribution.
- `T_Inv_2T(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the t-value of the Student t-distribution as a function of the probability and the degrees of freedom.
    - `Arg1 As Double` (required): Probability - The probability associated with the two-tailed Student t-distribution.
    - `Arg2 As Double` (required): Degrees_freedom - The number of degrees of freedom with which to characterize the distribution.
- `Var_S(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Estimates variance based on a sample.
- `Var_P(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Calculates variance based on the entire population.
- `Weibull_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean) As Double`  
  Returns the Weibull distribution. Use this distribution in reliability analysis, such as calculating the mean time to failure for a device.
    - `Arg1 As Double` (required): X - The value at which to evaluate the function.
    - `Arg2 As Double` (required): Alpha - A parameter to the distribution.
    - `Arg3 As Double` (required): Beta - A parameter to the distribution.
    - `Arg4 As Boolean` (required): Cumulative - Determines the form of the function.
- `NetworkDays_Intl(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant], [Arg4 As Variant]) As Double`  
  Returns the number of whole workdays between two dates using parameters to indicate which and how many days are weekend days. Weekend days and any days that are specified as holidays are not considered as workdays.
    - `Arg1 As Variant` (required): Start_date - The start date for which the difference is to be computed. The start_date can be earlier than, the same as, or later than the end_date.
    - `Arg2 As Variant` (required): End_date - The end date for which the difference is to be computed. The start_date can be earlier than, the same as, or later than the end_date.
    - `Arg3 As Variant` (optional): Weekend - Indicates the days of the week that are weekend days and are not included in the number of whole working days between start_date and end_date. Weekend is a weekend number or string that specifies when weekends occur. Weekend number values indicate the weekend days listed in the following table.<table><tr><th>Weekend number</th><th>Weekend days</th> </tr><tr><td>1 or omitted</td><td>Saturday, Sunday</td></tr><tr><td>2</td><td>Sunday, Monday</td></tr><tr><td>3</td><td>Monday, Tuesday </td> </tr><tr><td>4</td><td>Tuesday, Wednesday</td> </tr><tr><td>5</td><td>Wednesday, Thursday</td> </tr><tr><td>6</td><td>Thursday, Friday</td></tr><tr><td>7</td><td>Friday, Saturday</td></tr><tr><td>11</td><td>Sunday only</td></tr><tr><td>12</td><td>Monday only</td></tr><tr><td>13</td><td>Tuesday only</td> </tr><tr><td>14</td><td>Wednesday only</td></tr><tr><td>15</td><td>Thursday only</td> </tr><tr><td>16</td><td>Friday only</td></tr><tr><td>17</td><td>Saturday only</td> </tr></table>
    - `Arg4 As Variant` (optional): Holidays - An optional set of one or more dates that are to be excluded from the working day calendar. Holidays is a range of cells that contain the dates, or an array constant of the serial values that represent those dates. The ordering of dates or serial values in holidays can be arbitrary.
- `WorkDay_Intl(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant], [Arg4 As Variant]) As Double`  
  Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters. Weekend parameters indicate which and how many days are weekend days. Weekend days and any days that are specified as holidays are not considered as workdays.
    - `Arg1 As Variant` (required): Start_date - The start date, truncated to integer.
    - `Arg2 As Variant` (required): Days - The number of workdays before or after the start_date. A positive value yields a future date; a negative value yields a past date; a 0 (zero) value yields the start_date. Day-offset is truncated to an integer.
    - `Arg3 As Variant` (optional): Weekend - Indicates the days of the week that are weekend days and are not considered working days. Weekend is a weekend number or string that specifies when weekends occur. Weekend number values indicate the following weekend days.<table><tr><th>Weekend number</th><th>Weekend days</th> </tr><tr><td>1 or omitted</td><td>Saturday, Sunday</td></tr><tr><td>2</td><td>Sunday, Monday</td></tr><tr><td>3</td><td>Monday, Tuesday </td> </tr><tr><td>4</td><td>Tuesday, Wednesday</td> </tr><tr><td>5</td><td>Wednesday, Thursday</td> </tr><tr><td>6</td><td>Thursday, Friday</td></tr><tr><td>7</td><td>Friday, Saturday</td></tr><tr><td>11</td><td>Sunday only</td></tr><tr><td>12</td><td>Monday only</td></tr><tr><td>13</td><td>Tuesday only</td> </tr><tr><td>14</td><td>Wednesday only</td></tr><tr><td>15</td><td>Thursday only</td> </tr><tr><td>16</td><td>Friday only</td></tr><tr><td>17</td><td>Saturday only</td> </tr></table>
    - `Arg4 As Variant` (optional): Holidays - An optional set of one or more dates that are to be excluded from the working day calendar. Holidays is a range of cells that contain the dates, or an array constant of the serial values that represent those dates. The ordering of dates or serial values in holidays can be arbitrary.
- `ISO_Ceiling(Arg1 As Double, [Arg2 As Variant]) As Double`  
  Returns a number that is rounded up to the nearest integer or to the nearest multiple of significance.
    - `Arg1 As Double` (required): Number - The value to be rounded.
    - `Arg2 As Variant` (optional): Significance - The optional multiple to which number is to be rounded. If significance is omitted, its default value is 1. NOTE: The absolute value of the multiple is used, so that the ISO_Ceiling function returns the mathematical ceiling irrespective of the signs of number and significance.
- `Beta_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean, [Arg5 As Variant], [Arg6 As Variant]) As Double`  
  Returns the beta cumulative distribution function.
    - `Arg1 As Double` (required): The value between A and B at which to evaluate the function.
    - `Arg2 As Double` (required): The Alpha parameter of the distribution.
    - `Arg3 As Double` (required): The Beta parameter of the distribution.
    - `Arg4 As Boolean` (required): Cumulative - a logical value that determines the form of the function. If cumulative is True, BETA.DIST returns the cumulative distribution function; if False, it returns the probability density function.
    - `Arg5 As Variant` (optional): An optional lower bound to the interval of x.
    - `Arg6 As Variant` (optional): An optional upper bound to the interval of x.
- `Beta_Inv(Arg1 As Double, Arg2 As Double, Arg3 As Double, [Arg4 As Variant], [Arg5 As Variant]) As Double`  
  Returns the inverse of the cumulative distribution function for a specified beta distribution. That is, if probability = Beta_Dist(x,...), then Beta_Inv(probability,...) = x.
    - `Arg1 As Double` (required): A probability associated with the beta distribution.
    - `Arg2 As Double` (required): The Alpha parameter of the distribution.
    - `Arg3 As Double` (required): The Beta parameter the distribution.
    - `Arg4 As Variant` (optional): An optional lower bound to the interval of x.
    - `Arg5 As Variant` (optional): An optional upper bound to the interval of x.
- `ChiSq_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Boolean) As Double`  
  Returns the chi-squared distribution.
    - `Arg1 As Double` (required): X - The value at which you want to evaluate the distribution.
    - `Arg2 As Double` (required): Deg_freedom - The number of degrees of freedom.
    - `Arg3 As Boolean` (required): Cumulative - A logical value that determines the form of the function. If cumulative is True, ChiSq_Dist returns the cumulative distribution function; if False, it returns the probability density function.
- `ChiSq_Dist_RT(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the right-tailed probability of the chi-squared distribution.
    - `Arg1 As Double` (required): The value at which you want to evaluate the distribution.
    - `Arg2 As Double` (required): The number of degrees of freedom.
- `ChiSq_Inv(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the inverse of the left-tailed probability of the chi-squared distribution.
    - `Arg1 As Double` (required): Probability - A probability associated with the chi-squared distribution.
    - `Arg2 As Double` (required): Deg_freedom - The number of degrees of freedom.
- `ChiSq_Inv_RT(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the inverse of the right-tailed probability of the chi-squared distribution.
    - `Arg1 As Double` (required): A probability associated with the chi-squared distribution.
    - `Arg2 As Double` (required): The number of degrees of freedom.
- `F_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean) As Double`  
  Returns the F probability distribution.
    - `Arg1 As Double` (required): X - The value at which to evaluate the function.
    - `Arg2 As Double` (required): Deg_freedom1 - The numerator degrees of freedom.
    - `Arg3 As Double` (required): Deg_freedom2 - The denominator degrees of freedom.
    - `Arg4 As Boolean` (required): Cumulative - A logical value that determines the form of the function. If cumulative is True, F_DIST returns the cumulative distribution function; if False, it returns the probability density function.
- `F_Dist_RT(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the right-tailed F probability distribution. Use this function to determine whether two data sets have different degrees of diversity. For example, you can examine the test scores of men and women entering high school and determine if the variability in the females is different from that found in the males.
    - `Arg1 As Double` (required): X - the value at which to evaluate the function.
    - `Arg2 As Double` (required): Degrees_freedom1 - the numerator degrees of freedom.
    - `Arg3 As Double` (required): Degrees_freedom2 - the denominator degrees of freedom.
- `F_Inv(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the inverse of the F probability distribution.
    - `Arg1 As Double` (required): Probability - A probability associated with the F cumulative distribution.
    - `Arg2 As Double` (required): Deg_freedom1 - The numerator degrees of freedom.
    - `Arg3 As Double` (required): Deg_freedom2 - The denominator degrees of freedom.
- `F_Inv_RT(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the inverse of the right-tailed F probability distribution. If p = F_DIST_RT(x,...), then F_INV_RT(p,...) = x.
    - `Arg1 As Double` (required): Probability - a probability associated with the F cumulative distribution.
    - `Arg2 As Double` (required): Degrees_freedom1 - the numerator degrees of freedom.
    - `Arg3 As Double` (required): Degrees_freedom2 - the denominator degrees of freedom.
- `HypGeom_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Double, Arg5 As Boolean) As Double`  
  Returns the hypergeometric distribution. HypGeom_Dist returns the probability of a given number of sample successes, given the sample size, population successes, and population size. Use HypGeom_Dist for problems with a finite population, where each observation is either a success or a failure, and where each subset of a given size is chosen with equal likelihood.
    - `Arg1 As Double` (required): Sample_s - the number of successes in the sample.
    - `Arg2 As Double` (required): Number_sample - the size of the sample.
    - `Arg3 As Double` (required): Population_s - the number of successes in the population.
    - `Arg4 As Double` (required): Number_population - the population size.
    - `Arg5 As Boolean` (required): Cumulative - a logical value that determines the form of the function. If cumulative is True, HypGeom_Dist returns the cumulative distribution function; if False, it returns the probability mass function.
- `LogNorm_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean) As Double`  
  Returns the lognormal distribution of x, where ln(x) is normally distributed with parameters mean and standard_dev. Use this function to analyze data that has been logarithmically transformed.
    - `Arg1 As Double` (required): X - The value at which to evaluate the function.
    - `Arg2 As Double` (required): Mean - The mean of ln(x).
    - `Arg3 As Double` (required): Standard_dev - The standard deviation of ln(x).
    - `Arg4 As Boolean` (required): Cumulative - A logical value that determines the form of the function. If cumulative is True, LogNorm_Dist returns the cumulative distribution function; if False, it returns the probability density function.
- `LogNorm_Inv(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the inverse of the lognormal cumulative distribution function. Use the lognormal distribution to analyze logarithmically transformed data.
    - `Arg1 As Double` (required): Probability - A probability associated with the lognormal distribution.
    - `Arg2 As Double` (required): Mean - The mean of ln(x).
    - `Arg3 As Double` (required): Standard_dev - The standard deviation of ln(x).
- `NegBinom_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean) As Double`  
  Returns the negative binomial distribution. NegBinom_Dist returns the probability that there will be number_f failures before the number_s-th success, when the constant probability of a success is probability_s. This function is similar to the binomial distribution, except that the number of successes is fixed, and the number of trials is variable. Like the binomial, trials are assumed to be independent.
    - `Arg1 As Double` (required): Number_f - the number of failures.
    - `Arg2 As Double` (required): Number_s - the threshold number of successes.
    - `Arg3 As Double` (required): Probability_s - the probability of a success.
    - `Arg4 As Boolean` (required): Cumulative - A logical value that determines the form of the function. If cumulative is True, NegBinom_Dist returns the cumulative distribution function; if False, it returns the probability density function.
- `Norm_S_Dist(Arg1 As Double, Arg2 As Boolean) As Double`  
  Returns the standard normal cumulative distribution function. The distribution has a mean of 0 (zero) and a standard deviation of one. Use this function in place of a table of standard normal curve areas.
    - `Arg1 As Double` (required): Z - The value for which you want the distribution.
    - `Arg2 As Boolean` (required): Cumulative - A logical value that determines the form of the function. If cumulative is True, Norm_S_Dist returns the cumulative distribution function; if False, it returns the probability mass function.
- `Norm_S_Inv(Arg1 As Double) As Double`  
  Returns the inverse of the standard normal cumulative distribution. The distribution has a mean of 0 (zero) and a standard deviation of one.
    - `Arg1 As Double` (required): Probability - A probability corresponding to the normal distribution.
- `T_Test(Arg1 As Variant, Arg2 As Variant, Arg3 As Double, Arg4 As Double) As Double`  
  Returns the probability associated with a Student t-Test. Use T_Test to determine whether two samples are likely to have come from the same two underlying populations that have the same mean.
    - `Arg1 As Variant` (required): Array1 - The first data set.
    - `Arg2 As Variant` (required): Array2 - The second data set.
    - `Arg3 As Double` (required): Tails - Specifies the number of distribution tails. If tails = 1, T_Test uses the one-tailed distribution. If tails = 2, T_Test uses the two-tailed distribution.
    - `Arg4 As Double` (required): Type - The kind of t-Test to perform.
- `Z_Test(Arg1 As Variant, Arg2 As Double, [Arg3 As Variant]) As Double`  
  Returns the one-tailed probability-value of a z-test. For a given hypothesized population mean, Z_Test returns the probability that the sample mean would be greater than the average of observations in the data set (_array_); that is, the observed sample mean.
    - `Arg1 As Variant` (required): _Array_ is the array or range of data against which to test the hypothesized population mean.
    - `Arg2 As Double` (required): The value to test.
    - `Arg3 As Variant` (optional): _Sigma_ is the population (known) standard deviation. If omitted, the sample standard deviation is used.
- `Binom_Dist(Arg1 As Double, Arg2 As Double, Arg3 As Double, Arg4 As Boolean) As Double`  
  Returns the individual term binomial distribution probability.
    - `Arg1 As Double` (required): Number_s - the number of successes in trials.
    - `Arg2 As Double` (required): Trials - the number of independent trials.
    - `Arg3 As Double` (required): Probability_s - the probability of success on each trial.
    - `Arg4 As Boolean` (required): Cumulative - a logical value that determines the form of the function. If cumulative is True, the Binom_Dist method returns the cumulative distribution function, which is the probability that there are at most number_s successes; if False, it returns the probability mass function, which is the probability that there are number_s successes.
- `Binom_Inv(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the inverse of the individual term binomial distribution probability.
    - `Arg1 As Double` (required): Trials - the number of Bernoulli trials.
    - `Arg2 As Double` (required): Probability_s - the probability of a success on each trial.
    - `Arg3 As Double` (required): Alpha - the criterion value.
- `Erf_Precise(Arg1 As Variant) As Double`  
  Returns the error function integrated between zero and lower_limit.
    - `Arg1 As Variant` (required): Lower_limit - the lower bound for integrating Erf.
- `ErfC_Precise(Arg1 As Variant) As Double`  
  Returns the complementary error function integrated between the specified value and infinity.
    - `Arg1 As Variant` (required): The value to compute against.
- `GammaLn_Precise(Arg1 As Double) As Double`  
  Returns the natural logarithm of the gamma function, (x).
    - `Arg1 As Double` (required): X - the value for which you want to calculate GammaLn.
- `Ceiling_Precise(Arg1 As Double, [Arg2 As Variant]) As Double`  
  Returns the specified number rounded to the nearest multiple of significance.
    - `Arg1 As Double` (required): Number - the value that you want to round.
    - `Arg2 As Variant` (optional): Significance - the multiple to which you want to round.
- `Floor_Precise(Arg1 As Double, [Arg2 As Variant]) As Double`  
  Rounds the specified number to the nearest multiple of significance.
    - `Arg1 As Double` (required): Number - the numeric value that you want to round.
    - `Arg2 As Variant` (optional): Significance - the multiple to which you want to round.
- `Acot(Arg1 As Double) As Double`  
  Returns the arccotangent of a number, in radians in the range 0 (zero) to pi.
    - `Arg1 As Double` (required): The cotangent of the angle that you want.
- `Acoth(Arg1 As Double) As Double`  
  Returns the inverse hyperbolic cotangent of a number.
    - `Arg1 As Double` (required): The hyperbolic cotangent of the angle that you want.
- `Cot(Arg1 As Double) As Double`  
  Returns the cotangent of an angle.
    - `Arg1 As Double` (required): The angle in radians for which you want the cotangent.
- `Coth(Arg1 As Double) As Double`  
  Returns the hyperbolic cotangent of a number.
    - `Arg1 As Double` (required): The angle in radians for which you want the hyperbolic cotangent.
- `Csc(Arg1 As Double) As Double`  
  Returns the cosecant of an angle.
    - `Arg1 As Double` (required): The angle in radians for which you want the cosecant.
- `Csch(Arg1 As Double) As Double`  
  Returns the hyperbolic cosecant of an angle.
    - `Arg1 As Double` (required): The angle in radians for which you want the hyperbolic cosecant.
- `Sec(Arg1 As Double) As Double`  
  Returns the secant of an angle.
    - `Arg1 As Double` (required): Angle in radians for which you want the secant.
- `Sech(Arg1 As Double) As Double`  
  Returns the hyperbolic secant of an angle.
    - `Arg1 As Double` (required): Angle in radians for which you want the hyperbolic secant.
- `ImCot(Arg1 As Variant) As String`  
  Returns the cotangent of a complex number.
    - `Arg1 As Variant` (required): Complex number for which you want the cotangent.
- `ImTan(Arg1 As Variant) As String`  
  Returns the tangent of a complex number.
    - `Arg1 As Variant` (required): Complex number for which you want the tangent.
- `ImCsc(Arg1 As Variant) As String`  
  Returns the cosecant of a complex number.
    - `Arg1 As Variant` (required): Complex number for which you want the cosecant.
- `ImCsch(Arg1 As Variant) As String`  
  Returns the hyperbolic cosecant of a complex number.
    - `Arg1 As Variant` (required): Complex number for which you want the hyperbolic cosecant.
- `ImSec(Arg1 As Variant) As String`  
  Returns the secant of a complex number.
    - `Arg1 As Variant` (required): Complex number for which you want the secant.
- `ImSech(Arg1 As Variant) As String`  
  Returns the hyperbolic secant of a complex number.
    - `Arg1 As Variant` (required): Complex number for which you want the hyperbolic secant.
- `Bitand(Arg1 As Double, Arg2 As Double) As Double`  
  Returns a bitwise And of two numbers.
    - `Arg1 As Double` (required): The decimal representation of the binary number that you want to evaluate.
    - `Arg2 As Double` (required): The decimal representation of the binary number that you want to evaluate.
- `Bitor(Arg1 As Double, Arg2 As Double) As Double`  
  Returns a bitwise Or of two numbers.
    - `Arg1 As Double` (required): The decimal representation of the binary number that you want to evaluate.
    - `Arg2 As Double` (required): The decimal representation of the binary number that you want to evaluate.
- `Bitxor(Arg1 As Double, Arg2 As Double) As Double`  
  Returns a bitwise Exclusive Or of two numbers.
    - `Arg1 As Double` (required): The decimal representation of the binary number that you want to evaluate.
    - `Arg2 As Double` (required): The decimal representation of the binary number that you want to evaluate.
- `Bitlshift(Arg1 As Double, Arg2 As Double) As Double`  
  Returns a value number shifted left by shift_amount bits.
    - `Arg1 As Double` (required): The decimal representation of the binary number that you want to evaluate.
    - `Arg2 As Double` (required): The number of bits that you want to shift _Arg1_ left by.
- `Bitrshift(Arg1 As Double, Arg2 As Double) As Double`  
  Returns a value number shifted right by shift_amount bits.
    - `Arg1 As Double` (required): The decimal representation of the binary number that you want to evaluate.
    - `Arg2 As Double` (required): The number of bits that you want to shift _Arg1_ right by.
- `Xor(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Boolean`  
  Returns a logical exclusive OR of all arguments.
    - `Arg1 As Variant` (required): 1 to 254 conditions that you want to test that can be either True or False; can be logical values, arrays, or references.
- `Combina(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the number of combinations with repetitions for a given number of items.
    - `Arg1 As Double` (required): The total number of items.
    - `Arg2 As Double` (required): The number of items in each combination.
- `Permutationa(Arg1 As Double, Arg2 As Double) As Double`  
  Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects.
    - `Arg1 As Double` (required): Total number of objects.
    - `Arg2 As Double` (required): Number of objects in each permutation.
- `PDuration(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns the number of periods required by an investment to reach a specified value.
    - `Arg1 As Double` (required): Interest rate per period.
    - `Arg2 As Double` (required): Present value of the investment.
    - `Arg3 As Double` (required): Desired future value of the investment.
- `Base(Arg1 As Double, Arg2 As Double, [Arg3 As Variant]) As String`  
  Converts a number into a text representation with the given radix (base).
    - `Arg1 As Double` (required): The number that you want to convert.
    - `Arg2 As Double` (required): The base radix that you want to convert the number into.
    - `Arg3 As Variant` (optional): The minimum length of the returned string. If omitted, leading zeros are not added.
- `Decimal(Arg1 As String, Arg2 As Double) As Double`  
  Converts a text representation of a number in a given base into a decimal number.
    - `Arg1 As String` (required): The number that you want to convert.
    - `Arg2 As Double` (required): The base radix of the number that you are converting.
- `Days(Arg1 As Variant, Arg2 As Variant) As Double`  
  Returns the number of days between two dates.
    - `Arg1 As Variant` (required): start_date(_Arg2_) and end_date (_Arg1_) are the two dates between which you want to know the number of days.
    - `Arg2 As Variant` (required): start_date(_Arg2_) and end_date (_Arg1_) are the two dates between which you want to know the number of days.
- `Binom_Dist_Range(Arg1 As Double, Arg2 As Double, Arg3 As Double, [Arg4 As Variant]) As Double`  
  Returns the probability of a trial result using a binomial distribution.
    - `Arg1 As Double` (required): The number of independent trials.
    - `Arg2 As Double` (required): The probability of success on each trial.
    - `Arg3 As Double` (required): The number of successes in trials.
    - `Arg4 As Variant` (optional): If provided, this function returns the probability that the number of successful trials shall lie between _Arg3_ and _Arg4_.
- `Gamma(Arg1 As Double) As Double`  
  Returns the gamma function value.
    - `Arg1 As Double` (required): The value for which you want to calculate gamma.
- `Gauss(Arg1 As Double) As Double`  
  Returns 0.5 less than the standard normal cumulative distribution.
    - `Arg1 As Double` (required): The value for which you want the distribution.
- `Phi(Arg1 As Double) As Double`  
  Returns the value of the density function for a standard normal distribution.
    - `Arg1 As Double` (required): Number for which you want the density of the standard normal distribution.
- `Skew_p(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Double`  
  Returns the skewness of a distribution based on a population: a characterization of the degree of asymmetry of a distribution around its mean.
    - `Arg1 As Variant` (required): 1 to 254 numbers or names, arrays, or references that contain numbers for which you want the population skewness.
- `Rri(Arg1 As Double, Arg2 As Double, Arg3 As Double) As Double`  
  Returns an equivalent interest rate for the growth of an investment.
    - `Arg1 As Double` (required): Number of periods for the investment.
    - `Arg2 As Double` (required): Present value of the investment.
    - `Arg3 As Double` (required): Future value of the investment.
- `Unichar(Arg1 As Double) As String`  
  Returns the Unicode character referenced by the given numeric value.
    - `Arg1 As Double` (required): Unicode number representing a character.
- `Unicode(Arg1 As String) As Double`  
  Returns the number (code point) corresponding to the first character of the text.
    - `Arg1 As String` (required): Character that you want the Unicode value of.
- `Munit(Arg1 As Double) As Variant`  
  Returns the unit matrix for the specified dimension.
    - `Arg1 As Double` (required): Integer specifying the dimension of the unit matrix that you want to return.
- `Arabic(Arg1 As String) As Double`  
  Converts a Roman numeral to an Arabic numeral.
    - `Arg1 As String` (required): The Roman numeral that you want to convert.
- `IsoWeekNum(Arg1 As Double, [Arg2 As Variant]) As Double`  
  Returns the ISO week number of the year for a given date.
    - `Arg1 As Double` (required): Date-time code used by Microsoft Excel for date and time calculation.
    - `Arg2 As Variant` (optional): This argument is not available in the function.
- `NumberValue(Arg1 As String, Arg2 As String, Arg3 As String) As Double`  
  Converts text to number in a locale-independent manner.
    - `Arg1 As String` (required): String representing the number that you want to convert.
    - `Arg2 As String` (required): Character used as the decimal separator in the string.
    - `Arg3 As String` (required): Character used as the group separator in the string.
- `IsFormula(Arg1 As Range) As Boolean`  
  Checks whether a reference is to a cell containing a formula, and returns True or False.
    - `Arg1 As Range` (required): Reference to the cell that you want to test. Reference can be a cell reference, a formula, or a name that refers to a cell.
- `IfNa(Arg1 As Variant, Arg2 As Variant) As Variant`  
  Returns the value that you specify if the expression resolves to #N/A; otherwise, returns the result of the expression.
    - `Arg1 As Variant` (required): value: Any value or expression or reference.
    - `Arg2 As Variant` (required): value_if_na: Any value or expression or reference.
- `Ceiling_Math(Arg1 As Double, [Arg2 As Variant], [Arg3 As Variant]) As Double`  
  Rounds a number up to the nearest integer or to the nearest multiple of significance.
    - `Arg1 As Double` (required): The value that you want to round.
    - `Arg2 As Variant` (optional): The multiple to which you want to round.
    - `Arg3 As Variant` (optional): When given and nonzero, this function will round away from zero.
- `Floor_Math(Arg1 As Double, [Arg2 As Variant], [Arg3 As Variant]) As Double`  
  Rounds a number down, to the nearest integer or to the nearest multiple of significance.
    - `Arg1 As Double` (required): The value that you want to round.
    - `Arg2 As Variant` (optional): The multiple to which you want to round.
    - `Arg3 As Variant` (optional): When given a nonzero number, this function will round towards zero.
- `ImSinh(Arg1 As Variant) As String`  
  Returns the hyperbolic sine of a complex number.
    - `Arg1 As Variant` (required): Complex number for which you want the hyperbolic sine.
- `ImCosh(Arg1 As Variant) As String`  
  Returns the hyperbolic cosine of a complex number.
    - `Arg1 As Variant` (required): Complex number for which you want the hyperbolic cosine.
- `FilterXML(Arg1 As String, Arg2 As String) As Variant`  
  Gets specific data from the returned XML, typically from a WebService function call.
    - `Arg1 As String` (required): Valid XML string.
    - `Arg2 As String` (required): XPath query string.
- `WebService(Arg1 As String) As Variant`  
  Underlying function that calls the web service asynchronously, using an HTTP GET request, and returns the response.
- `EncodeURL(Arg1 As String) As Variant`  
  URL encodes the argument.
    - `Arg1 As String` (required): Text to be encoded.
- `Forecast_ETS(Arg1 As Double, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant]) As Double`  
  Calculates or predicts a future value based on existing (historical) values by using the AAA version of the Exponential Smoothing (ETS) algorithm.
    - `Arg1 As Double` (required): Target Date: the data point for which you want to predict a value. Target date can be date/time or numeric. See Remarks.
    - `Arg2 As Variant` (required): Values: the historical values, for which you want to forecast the next points.
    - `Arg3 As Variant` (required): Timeline: the independent array or range of dates or numeric data. The values in the timeline must have a consistent step between them and can't be zero. See Remarks.
    - `Arg4 As Variant` (optional): Seasonality: A numeric value. See Remarks.
    - `Arg5 As Variant` (optional): Data completions: Although the timeline requires a constant step between data points, Forecast_ETS supports up to 30% missing data, and automatically adjusts for it. See Remarks.
    - `Arg6 As Variant` (optional): Aggregation: Although the timeline requires a constant step between data points, Forecast_ETS aggregates multiple points that have the same time stamp. See Remarks.
- `Forecast_ETS_ConfInt(Arg1 As Double, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant]) As Double`  
  Returns a confidence interval for the forecast value at the specified target date.
    - `Arg1 As Double` (required): Target Date: the data point for which you want to predict a value. Target date can be date/time or numeric. See Remarks.
    - `Arg2 As Variant` (required): Values: the historical values, for which you want to forecast the next points.
    - `Arg3 As Variant` (required): Timeline: the independent array or range of dates or numeric data. The values in the timeline must have a consistent step between them and can't be zero. See Remarks.
    - `Arg4 As Variant` (optional): Confidence level: A numerical value between 0 and 1 (exclusive), indicating a confidence level for the calculated confidence interval. See Remarks.
    - `Arg5 As Variant` (optional): Seasonality: A numeric value. See Remarks.
    - `Arg6 As Variant` (optional): Data completions: Although the timeline requires a constant step between data points, Forecast_ETS_ConfInt supports up to 30% missing data, and automatically adjusts for it. See Remarks.
    - `Arg7 As Variant` (optional): Aggregation: Although the timeline requires a constant step between data points, Forecast_ETS_ConfInt aggregates multiple points that have the same time stamp. See Remarks.
- `Forecast_ETS_Seasonality(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant], [Arg4 As Variant]) As Double`  
  Returns the length of the repetitive pattern that Excel detects for the specified time series.
    - `Arg1 As Variant` (required): Values: the historical values, for which you want to forecast the next points.
    - `Arg2 As Variant` (required): Timeline: the independent array or range of dates or numeric data. The values in the timeline must have a consistent step between them and can't be zero. See Remarks.
    - `Arg3 As Variant` (optional): Data completions: Although the timeline requires a constant step between data points, Forecast_ETS_Seasonality supports up to 30% missing data, and automatically adjusts for it. See Remarks.
    - `Arg4 As Variant` (optional): Aggregation: Although the timeline requires a constant step between data points, Forecast_ETS_Seasonality aggregates multiple points that have the same time stamp. See Remarks.
- `Forecast_Linear(Arg1 As Double, Arg2 As Variant, Arg3 As Variant) As Double`  
  Calculates, or predicts, a future value by using existing values. The predicted value is a y-value for a given x-value. The known values are existing x-values and y-values, and the new value is predicted by using linear regression. Use this function to predict future sales, inventory requirements, or consumer trends.
    - `Arg1 As Double` (required): x - the data point for which you want to predict a value.
    - `Arg2 As Variant` (required): known_y's - the dependent array or range of data.
    - `Arg3 As Variant` (required): known_x's - the independent array or range of data.
- `Forecast_ETS_STAT(Arg1 As Variant, Arg2 As Variant, Arg3 As Double, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant]) As Double`  
  Returns a statistical value as a result of time series forecasting.
    - `Arg1 As Variant` (required): Values: the historical values, for which you want to forecast the next points.
    - `Arg2 As Variant` (required): Timeline: the independent array or range of dates or numeric data. The values in the timeline must have a consistent step between them and can't be zero. See Remarks.
    - `Arg3 As Double` (required): Statistic_type: A numeric value between 1 and 8, indicating which statistic will be returned for the calculated forecast.
    - `Arg4 As Variant` (optional): Confidence level: A numerical value between 0 and 1 (exclusive), indicating a confidence level for the calculated confidence interval. See Remarks.
    - `Arg5 As Variant` (optional): Data completions: Although the timeline requires a constant step between data points, Forecast_ETS_STAT supports up to 30% missing data, and automatically adjusts for it. See Remarks.
    - `Arg6 As Variant` (optional): Aggregation: Although the timeline requires a constant step between data points, Forecast_ETS_STAT aggregates multiple points that have the same time stamp. See Remarks.
- `MaxIfs(Arg1 As Range, Arg2 As Range, Arg3 As Variant, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant]) As Double`
- `MinIfs(Arg1 As Range, Arg2 As Range, Arg3 As Variant, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant]) As Double`
- `TextJoin(Arg1 As String, Arg2 As Boolean, Arg3 As String, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant]) As String`
- `Concat(Arg1 As String, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant]) As String`
- `Sort(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant]) As Variant`
- `Unique(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant]) As Variant`
- `FieldValue(Arg1 As Variant, Arg2 As String) As Variant`
- `Filter(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant]) As Variant`
- `Sequence(Arg1 As Variant, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant]) As Variant`
- `RandArray([Arg1 As Variant], [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant]) As Variant`
- `Single(Arg1 As Variant) As Variant`
- `SortBy(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Variant`
- `XMatch(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant], [Arg4 As Variant]) As Double`
- `XLookup(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant]) As Variant`
- `StockHistory(Arg1 As Variant, Arg2 As Variant, [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant]) As Variant`
- `ValueToText(Arg1 As Variant, [Arg2 As Variant]) As String`
- `ArrayToText(Arg1 As Variant, [Arg2 As Variant]) As String`
