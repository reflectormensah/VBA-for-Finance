Attribute VB_Name = "Functions"
'Declaring the function
Function ForextoUSD(AmountInEUR, InputCurrency)

'Automatically updating the excel sheet when there a change
Application.Volatile True

RateRangeName = "USDper" & InputCurrency

'Formula for converting the function from EUR to USD
ForexRate = ws_Visual.Range(RateRangeName).Value

AmountInUSD = AmountInEUR * ForexRate


ForextoUSD = AmountInUSD
End Function

