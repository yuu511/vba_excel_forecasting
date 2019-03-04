
n array_len(arr As Variant) As Integer
  array_len = UBound(arr) - LBound(arr) + 1
End Function

' Convert row,column number into equivalent excel cell
Function parse_cell(row As Integer, col As Integer) As String
    Dim col_letter As String
    row_number = CStr(row)
    col_letter = Split(Cells(1, col).Address, "$")(1)
    parse_cell = col_letter + row_number
End Function

' Static forecasting
Function static_forecast(periodicity As Integer) As Double
    Dim col_n As Variant
    Dim row As Integer
    Dim col As Integer
    
    col_n = Array("Year", "Period", "Demand", "De-seasonalized Demand", "Regressed,Deseasonalized Demand", "Seasonal Factors", "Average Seasonal Demand", "Reintroduce Seasonal factors", "MAD", "Percent error", "MAPE", "TS")
    row = 1
    col = 1
    
    Dim i
    For i = 0 To array_len(col_n) - 1
      Worksheets("Sheet1").Range(parse_cell(row, col)).Value = col_n(i)
      col = col + 1
    Next i
End Function


Sub main()
  Dim p As Integer
  p = 4
  static_fcast = static_forecast(p)
End Sub

