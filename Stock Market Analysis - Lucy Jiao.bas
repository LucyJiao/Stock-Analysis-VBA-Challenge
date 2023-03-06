Attribute VB_Name = "Module1"

Sub StockAnalysis():

    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call StockAnalysis
    Next
    Application.ScreenUpdating = True
    
End Sub

Sub StockAnalysis():

Cells(1, 9) = "Ticker_Symbol"
Cells(1, 10) = "Yearly_Change"
Cells(1, 11) = "Percent_Change"
Cells(1, 12) = "Total_Stock_Volume"

Cells(1, 9).Font.Bold = True
Cells(1, 10).Font.Bold = True
Cells(1, 11).Font.Bold = True
Cells(1, 12).Font.Bold = True


Columns(9).AutoFit
Columns(10).AutoFit
Columns(11).AutoFit
Columns(12).AutoFit
Columns(12).AutoFit
    
Dim TICKER As Integer
TICKER = 1

Dim SummaryTableRow As Long
SummaryTableRow = 2

Dim TotalVolume  As Double
TotalVolume = 0

Dim OpeningValue As Double

Dim ClosingValue As Double

Dim ClosingMinusOpening As Double


rowmax = Cells(Rows.Count, "A").End(xlUp).Row


 For I = 2 To rowmax


    
    TotalVolume = TotalVolume + Cells(I, 7).Value
           
    If Cells(I - 1, TICKER).Value <> Cells(I, TICKER).Value Then
    OpeningValue = Cells(I, 3).Value
    
    End If

   
    If Cells(I + 1, TICKER).Value <> Cells(I, TICKER).Value Then

    ClosingValue = Cells(I, 6).Value
    
    ClosingMinusOpening = ClosingValue - OpeningValue
    
    Cells(SummaryTableRow, 9).Value = Cells(I, TICKER).Value
    
    Cells(SummaryTableRow, 10).Value = ClosingMinusOpening
    
    Cells(SummaryTableRow, 11).Value = ClosingMinusOpening / OpeningValue
    
    Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
    
    Cells(SummaryTableRow, 12).Value = TotalVolume
    
    SummaryTableRow = SummaryTableRow + 1

    TotalVolume = 0


    End If
    
Next I

   Columns("K:K").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
    End With
    Selection.FormatConditions(1).StopIfTrue = False
  
     Columns("J:J").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
    End With
    Selection.FormatConditions(1).StopIfTrue = False
 
    
Range("K1").Select
    Selection.FormatConditions.Delete
    
Range("J1").Select
    Selection.FormatConditions.Delete

Cells(1, 15) = "Ticker"
Cells(1, 16) = "Value"
Cells(2, 14) = "Greatest % Increase"
Cells(3, 14) = "Greatest % Decrease"
Cells(4, 14) = "Greatest Total Volume"


Columns(14).AutoFit


Cells(1, 15).Font.Bold = True
Cells(1, 16).Font.Bold = True
Cells(2, 14).Font.Bold = True
Cells(3, 14).Font.Bold = True
Cells(4, 14).Font.Bold = True

'-----------------------------------------------------------------------

Dim MaxValue As Double
Dim MinValue As Double
Dim GreatestTotalVolume As Double


    MaxValue = Application.WorksheetFunction.Max(Range("K:k"))

        Cells(2, 16) = MaxValue
        Cells(2, 16).NumberFormat = "0.00%"

    MinValue = Application.WorksheetFunction.Min(Range("K:k"))

        Cells(3, 16) = MinValue
        Cells(3, 16).NumberFormat = "0.00%"

    GreatestTotalVolume = Application.WorksheetFunction.Max(Range("L:l"))

        Cells(4, 16) = GreatestTotalVolume

Dim inc_loc As Integer
Dim dec_loc As Integer
Dim totalvolloc As Integer

inc_loc = WorksheetFunction.Match(MaxValue, Range("K:K"), 0)
dec_loc = WorksheetFunction.Match(MinValue, Range("K:K"), 0)
totalvolloc = WorksheetFunction.Match(GreatestTotalVolume, Range("L:L"), 0)

        Range("O2") = Cells(inc_loc + 1, 9)
        Range("O3") = Cells(dec_loc + 1, 9)
        Range("O4") = Cells(totalvolloc + 1, 9)



End Sub
