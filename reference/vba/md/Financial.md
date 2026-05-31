# Financial

**Type:** Module  
**Library:** Visual Basic For Applications  

## Functions (13)

- `SLN(Cost As Double, Salvage As Double, Life As Double) As Double`  
  Returns a Double specifying the straight-line depreciation of an asset for a single period.
    - `Cost As Double` (required): Required. Double specifying initial cost of the asset.
    - `Salvage As Double` (required): Required. Double specifying value of the asset at the end of its useful life.
    - `Life As Double` (required): Required. Double specifying length of the useful life of the asset.
- `SYD(Cost As Double, Salvage As Double, Life As Double, Period As Double) As Double`  
  Returns a Double specifying the sum-of-years' digits depreciation of an asset for a specified period.
    - `Cost As Double` (required): Required. Double specifying initial cost of the asset.
    - `Salvage As Double` (required): Required. Double specifying value of the asset at the end of its useful life.
    - `Life As Double` (required): Required. Double specifying length of the useful life of the asset.
    - `Period As Double` (required): Required. Double specifying period for which asset depreciation is calculated.
- `DDB(Cost As Double, Salvage As Double, Life As Double, Period As Double, [Factor As Variant]) As Double`  
  Returns a Double specifying the depreciation of an asset for a specific time period by using the double-declining balance method or some other method you specify.
    - `Cost As Double` (required): Required. Double specifying the initial cost of the asset.
    - `Salvage As Double` (required): Required. Double specifying the value of the asset at the end of its useful life.
    - `Life As Double` (required): Required. Double specifying the length of useful life of the asset.
    - `Period As Double` (required): Required. Double specifying the period for which asset depreciation is calculated.
    - `Factor As Variant` (optional): Optional. Variant specifying the rate at which the balance declines. If omitted, 2 (double-declining method) is assumed.
- `IPmt(Rate As Double, Per As Double, NPer As Double, PV As Double, [FV As Variant], [Due As Variant]) As Double`  
  Returns a Double specifying the interest payment for a given period of an annuity based on periodic, fixed payments and a fixed interest rate.
    - `Rate As Double` (required): Required. Double specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.
    - `Per As Double` (required): Required. Double specifying payment period in the range 1 through _nper_.
    - `NPer As Double` (required): Required. Double specifying total number of payment periods in the annuity. For example, if you make monthly payments on a four-year car loan, your loan has a total of 4 * 12 (or 48) payment periods.
    - `PV As Double` (required): Required. Double specifying present value, or value today, of a series of future payments or receipts. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you'll make.
    - `FV As Variant` (optional): Optional. Variant specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, $50,000 is the future value. If omitted, 0 is assumed.
- `PPmt(Rate As Double, Per As Double, NPer As Double, PV As Double, [FV As Variant], [Due As Variant]) As Double`  
  Returns a Double specifying the principal payment for a given period of an annuity based on periodic, fixed payments and a fixed interest rate.
    - `Rate As Double` (required): Required. Double specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.
    - `Per As Double` (required): Required. Integer specifying payment period in the range 1 through _nper_.
    - `NPer As Double` (required): Required. Integer specifying total number of payment periods in the annuity. For example, if you make monthly payments on a four-year car loan, your loan has a total of 4 * 12 (or 48) payment periods.
    - `PV As Double` (required): Required. Double specifying present value, or value today, of a series of future payments or receipts. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you'll make.
    - `FV As Variant` (optional): Optional. Variant specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, $50,000 is the future value. If omitted, 0 is assumed.
- `Pmt(Rate As Double, NPer As Double, PV As Double, [FV As Variant], [Due As Variant]) As Double`  
  Returns a Double specifying the payment for an annuity based on periodic, fixed payments and a fixed interest rate.
    - `Rate As Double` (required): Required. Double specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.
    - `NPer As Double` (required): Required. Integer specifying total number of payment periods in the annuity. For example, if you make monthly payments on a four-year car loan, your loan has a total of 4 * 12 (or 48) payment periods.
    - `PV As Double` (required): Required. Double specifying present value (or lump sum) that a series of payments to be paid in the future is worth now. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you'll make.
    - `FV As Variant` (optional): Optional. Variant specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, $50,000 is the future value. If omitted, 0 is assumed.
- `PV(Rate As Double, NPer As Double, Pmt As Double, [FV As Variant], [Due As Variant]) As Double`  
  Returns a Double specifying the present value of an annuity based on periodic, fixed payments to be paid in the future and a fixed interest rate.
    - `Rate As Double` (required): Required. Double specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.
    - `NPer As Double` (required): Required. Integer specifying total number of payment periods in the annuity. For example, if you make monthly payments on a four-year car loan, your loan has a total of 4 * 12 (or 48) payment periods.
    - `Pmt As Double` (required): Required. Double specifying payment to be made each period. Payments usually contain principal and interest that doesn't change over the life of the annuity.
    - `FV As Variant` (optional): Optional. Variant specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, $50,000 is the future value. If omitted, 0 is assumed.
- `FV(Rate As Double, NPer As Double, Pmt As Double, [PV As Variant], [Due As Variant]) As Double`  
  Returns a Double specifying the future value of an annuity based on periodic fixed payments and a fixed interest rate.
    - `Rate As Double` (required): Required. Double specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.
    - `NPer As Double` (required): Required. Integer specifying total number of payment periods in the annuity. For example, if you make monthly payments on a four-year car loan, your loan has a total of 4 * 12 (or 48) payment periods.
    - `Pmt As Double` (required): Required. Double specifying payment to be made each period. Payments usually contain principal and interest that doesn't change over the life of the annuity.
    - `PV As Variant` (optional): Optional. Variant specifying present value (or lump sum) of a series of future payments. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you'll make. If omitted, 0 is assumed.
- `NPer(Rate As Double, Pmt As Double, PV As Double, [FV As Variant], [Due As Variant]) As Double`  
  Returns a Double specifying the number of periods for an annuity based on periodic, fixed payments and a fixed interest rate.
    - `Rate As Double` (required): Required. Double specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.
    - `Pmt As Double` (required): Required. Double specifying payment to be made each period. Payments usually contain principal and interest that doesn't change over the life of the annuity.
    - `PV As Double` (required): Required. Double specifying present value, or value today, of a series of future payments or receipts. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you'll make.
    - `FV As Variant` (optional): Optional. Variant specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, $50,000 is the future value. If omitted, 0 is assumed.
- `Rate(NPer As Double, Pmt As Double, PV As Double, [FV As Variant], [Due As Variant], [Guess As Variant]) As Double`  
  Returns a Double specifying the interest rate per period for an annuity.
    - `NPer As Double` (required): Required. Double specifying total number of payment periods in the annuity. For example, if you make monthly payments on a four-year car loan, your loan has a total of 4 * 12 (or 48) payment periods.
    - `Pmt As Double` (required): Required. Double specifying payment to be made each period. Payments usually contain principal and interest that doesn't change over the life of the annuity.
    - `PV As Double` (required): Required. Double specifying present value, or value today, of a series of future payments or receipts. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you'll make.
    - `FV As Variant` (optional): Optional. Variant specifying future value or cash balance you want after you make the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, $50,000 is the future value. If omitted, 0 is assumed.
    - `Guess As Variant` (optional): Optional. Variant specifying value you estimate will be returned by Rate. If omitted, _guess_ is 0.1 (10 percent).
- `IRR(ValueArray As SAFEARRAY(Double), [Guess As Variant]) As Double`  
  Returns a Double specifying the internal rate of return for a series of periodic cash flows (payments and receipts).
    - `Guess As Variant` (optional): Optional. Variant specifying value that you estimate will be returned by IRR. If omitted, _guess_ is 0.1 (10 percent).
- `MIRR(ValueArray As SAFEARRAY(Double), FinanceRate As Double, ReinvestRate As Double) As Double`  
  Returns a Double specifying the modified internal rate of return for a series of periodic cash flows (payments and receipts).
- `NPV(Rate As Double, ValueArray As SAFEARRAY(Double)) As Double`  
  Returns a Double specifying the net present value of an investment based on a series of periodic cash flows (payments and receipts) and a discount rate.
    - `Rate As Double` (required): Required. Double specifying discount rate over the length of the period, expressed as a decimal.
