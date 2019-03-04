alculate number of demand datapoints starting at row,col
Function demand_len(row As Integer, col As Integer) As Integer
    Dim cnt As Integer
    cnt = 0
    While IsNumeric(Range(parse_cell(row, col))) = True And IsEmpty(Range(parse_cell(row, col))) = False
        cnt = cnt + 1
        row = row + 1
    Wend
    demand_len = cnt
End Function

' Get length of Array
Function array_len(arr As Variant) As Integer
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
    Dim num_demand As Integer
    Dim row As Integer
    Dim col As Integer
    Dim slope As Integer
    Dim intercept As Integer
    Set Wf = WorksheetFunction


    ' Calculate number of data points
    num_demand = demand_len(2, 3)
    
    ' Fill out column names
    row = 1
    col = 1
    col_n = Array("Year", "Period", "Demand", "De-seasonalized Demand", "Regressed,Deseasonalized Demand", "Seasonal Factors", "Average Seasonal Demand", "Reintroduce Seasonal factors", "Error", "Absolute Error", "Mean squared error", "MAD", "Percent error", "MAPE", "TS")
    Dim i
    For i = 0 To array_len(col_n) - 1
      Range(parse_cell(row, col)).Value = col_n(i)
      col = col + 1
    Next i
    
    ' Fill out year,period
    row = 2
    col = 1
    For i = 0 To num_demand - 1
      If i Mod periodicity = 0 Then
        Range(parse_cell(row, col)).Value = (i / periodicity) + 1
      End If
      Range(parse_cell(row, col + 1)).Value = i + 1
      row = row + 1
    Next i
    
   ' Deseasonalize Demand
   row = 2
   col = 4
   Dim min_val As Integer
   Dim max_val As Integer
   If periodicity Mod 2 = 0 Then
     min_val = (periodicity / 2) + 1
     max_val = num_demand - (periodicity / 2)
     For i = 0 To max_val - min_val
     Range(parse_cell(min_val + i + 1, col)).Formula = "=(" + parse_cell(row + i, 3) + "+" + parse_cell(row + periodicity + i, 3) + "+2*SUM(" + parse_cell(row + i + 1, 3) + ":" + parse_cell(row + periodicity + i - 1, 3) + "))/" + CStr(2 * periodicity)
     Next i
   Else
     min_val = (periodicity - 1) / 2 + 1
     max_val = num_demand - ((periodicity - 1) / 2)
     For i = 0 To max_val - min_val
     Range(parse_cell(min_val + i + 1, col)).Formula = "=(SUM(" + parse_cell(min_val + i, 3) + ":" + parse_cell(min_val + i + (periodicity - 1), 3) + ")" + "/" + CStr(periodicity) + ")"
     Next i
   End If
   
' Linearly regress the dataset and get the slope, intercept of the equation
Dim rX As String
Dim rY As String
rX = (parse_cell(min_val + 1, 2) + ":" + parse_cell(max_val + 1, 2))
rY = (parse_cell(min_val + 1, 4) + ":" + parse_cell(max_val + 1, 4))
slope = Wf.slope(Range(rY), Range(rX))
intercept = Wf.intercept(Range(rY), Range(rX))
Debug.Print (slope)
Debug.Print (intercept)
End Function



Sub main()
  Dim p As Integer
  p = 4
  static_fcast = static_forecast(p)
End Sub

