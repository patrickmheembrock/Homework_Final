Attribute VB_Name = "Module1"
Sub Mod_Homework()

For Each ws In Worksheets
ws.Activate

'Establish Variables
Dim Ticker As String
Dim YearBeg, YearEnd, Volume As Double
Dim Summary_Table_Row As Integer

'Create Summary Table Headers
Summary_Table_Row = 2
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Price Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Volume"
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For r = 2 To LastRow

'Find and Save your Opening Balance
If Cells(r, 1).Value <> Cells(r - 1, 1).Value Then
    YearBeg = Cells(r, 3).Value
End If

    'If the current ticker does not match the next ticker..
    If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
        
        'Assigns ticker name to be shown in the summary
        Ticker = Cells(r, 1).Value
        
        'Find and Save your Closing Balance
        YearEnd = Cells(r, 6).Value
        
        'Determine Final Volume
        Volume = Volume + Cells(r, 7).Value
                
        'Print Tickers in Summary Line
        Range("I" & Summary_Table_Row).Value = Ticker
        
        'Calculate and Print Dollar & Percentage Changes
        YearChg = YearEnd - YearBeg
        Range("J" & Summary_Table_Row).Value = YearChg
        Range("J" & Summary_Table_Row).NumberFormat = "0.00"
                        
            'Conditional Formatting for Yearly Change
            If YearChg >= 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
        'Calculate, print, and format Yearly % Change
        If YearBeg <> 0 Then
            Range("K" & Summary_Table_Row).Value = YearChg / YearBeg
        Else
            Range("K" & Summary_Table_Row).Value = "0"
        End If
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        'Print volume
        Range("L" & Summary_Table_Row).Value = Volume
        
        'Adjust/Reset the Variables
        Summary_Table_Row = Summary_Table_Row + 1
        YearBeg = 0
        Volume = 0
        
    'Or if the ticker is the same...
    Else
        '...Add the volume to the running total
        Volume = Volume + Cells(r, 7).Value
    End If
Next r


'Hard Portion
'Assign Additional Variables
Dim PerInc, PerDec, TotVol As Double
PerInc = Range("K2")
PerDec = Range("K2")
TotVol = Range("L2")

'Print Additional Summary Table
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"

'Find and Print Greatest % Increase
For r = 2 To LastRow
    If Cells(r, 11).Value > PerInc Then
        Ticker = Cells(r, 9).Value
        PerInc = Cells(r, 11).Value
    End If
Next r
Range("O2") = Ticker
Range("P2") = PerInc

'Find and Print Greatest % Decrease
For r = 2 To LastRow
    If Cells(r, 11).Value < PerDec Then
        Ticker = Cells(r, 9).Value
        PerDec = Cells(r, 11).Value
    End If
Next r
Range("O3") = Ticker
Range("P3") = PerDec

'Find and Print Greatest Total Volume
For r = 2 To LastRow
    If Cells(r, 12).Value > TotVol Then
        Ticker = Cells(r, 9).Value
        TotVol = Cells(r, 12).Value
    End If
Next r
Range("O4") = Ticker
Range("P4") = TotVol
     
'Format Percentages and Autofit Columns
Range("P2:P3").NumberFormat = "0.00%"
Columns("A:P").AutoFit


Next ws


End Sub


