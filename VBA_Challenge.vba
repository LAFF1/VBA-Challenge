'This module summarizes daily stock transactions by sheet
'Allow user to select sheet to be processed and makes that sheet active
'Adds columns: Ticker, Yearly Change, Percent Change, Total Stock Volume
'Adds columns for Greatest: Description, Ticker and Value
'Formats columns for easy viewing


'Declare Variable for user input used in more than 1 sub
Dim WSheet As String

Sub Main()

'Prompt for User Selection
UserWS = InputBox("Enter Worksheet Name" & vbCrLf & "or" & _
        vbCrLf & "Enter All to Proces All Sheets in the Workbook")
        
'Loop through all sheets to find selected or process "all"
For I = 1 To ActiveWorkbook.Worksheets.Count
        currentWS = ActiveWorkbook.Worksheets(I).Name 'Get Current sheet name
        'Decide if current sheet should be summarized, correct for entry case
    If LCase(UserWS) = "all" Or LCase(UserWS) = LCase(currentWS) Then
        Worksheets(currentWS).Activate 'activate chosen sheet
        WSheet = currentWS 'set Global sheet for use in subs
        GetStockSummary ' call sub to summarize sheet
    End If
Next I
    
End Sub

Sub GetStockSummary():

'Set titles for results columns
SetTitles

'Declare variable to hold Ticker Symbol
Dim Ticker As String

'Declare variable to hold Opening Price and set initial value
Dim Opening As Double
Opening = Cells(2, 3).Value

'Declare variable to hold Closing Price
Dim Closing As Double

'Declare variable to hold Change in Price
Dim Change As Double

'Declare variable to hold Change in Price
Dim PerChange As Double

'Declare variable to hold Volume and set initial value
Dim Volume As Double
Volume = 0

'Declare variable to hold result table position value
Dim ResultRow As Integer
ResultRow = 2

'Declare variables for Greatest Bonus
Dim GreatIncTick As String  'Declare Greatest Increase Ticker
Dim GreatInc As Double      'Declare Greatest Increase
Dim GreatDecTick As String  'Declare Greatest Decrease Ticker
Dim GreatDec As Double      'Declare Greatest Decrease
Dim GreatVolTick As String  'Declare Greatest Volume Ticker
Dim GreatVol As Double      'Declare Greatest Volume


'Declare variable to hold source sheet number of rows and set number of rows
Dim Maxrow As Long

'Find the end of the worksheet
Range("A" & Rows.Count).End(xlUp).Select 'Go to end of sheet
Maxrow = ActiveCell.Row 'Get the row number

'Return to top of sheet to view the results
Range("A1").Select

  For I = 2 To Maxrow
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        'Set values
        Ticker = Cells(I, 1).Value
        Closing = Cells(I, 6).Value
        'Find value change value
        Change = Closing - Opening
        'Find percentage for gain or loss
        PerChange = Change / Opening
        'Add final volume number
        Volume = Volume + Cells(I, 7).Value
        'Enter results on Sheet
        Range("I" & ResultRow).Value = Ticker
        Range("J" & ResultRow).Value = Change
        'Set cell colors for Gains and Losses
         If Change >= 0 Then
        ' set winners to green
            Range("J" & ResultRow).Interior.Color = vbGreen
        Else
        'Set losers to red
            Range("J" & ResultRow).Interior.Color = vbRed
        End If
        'Enter valueS on sheet for current ticker
        Range("K" & ResultRow).Value = FormatPercent(PerChange, 2)
        Range("L" & ResultRow).Value = Volume
        'Test for greatest and set value when
        'it is greater than currently held value
        If PerChange > GreatInc Then
            GreatInc = PerChange
            GreatIncTick = Ticker
        ElseIf PerChange < GreatDec Then
            GreatDec = PerChange
            GreatDecTick = Ticker
        End If
        If Volume > GreatVol Then
            GreatVol = Volume
            GreatVolTick = Ticker
        End If
        'Reset values for next Ticker Symbol
        Volume = 0
        ResultRow = ResultRow + 1
        Opening = Cells(I + 1, 3).Value
    Else
        'Add current row volume to running total
        Volume = Volume + Cells(I, 7).Value
        
    End If
    
  Next I
  
'Set Final Values for Greatest
  Cells(2, 16) = GreatIncTick
  Cells(2, 17) = FormatPercent(GreatInc, 2)
  Cells(3, 16) = GreatDecTick
  Cells(3, 17) = FormatPercent(GreatDec, 2)
  Cells(4, 16) = GreatVolTick
  Cells(4, 17) = GreatVol
  
'Resize Columns for easy Viewing
  Worksheets(WSheet).Columns("A:Q").AutoFit
  Worksheets(WSheet).Columns("P:P").AutoFit
  Worksheets(WSheet).Range("O2:O4").Columns.AutoFit
  
'Display a box around the Greatest - for fun!
  Range("O1:Q4").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDash
  Range("O1:Q4").Borders(xlEdgeTop).LineStyle = XlLineStyle.xlDash
  Range("O1:Q4").Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlDash
  Range("O1:Q4").Borders(xlEdgeRight).LineStyle = XlLineStyle.xlDash
'Set color of box to Blue
  Range("O1:Q4").Borders(xlEdgeBottom).Color = vbBlue
  Range("O1:Q4").Borders(xlEdgeTop).Color = vbBlue
  Range("O1:Q4").Borders(xlEdgeLeft).Color = vbBlue
  Range("O1:Q4").Borders(xlEdgeRight).Color = vbBlue

  
End Sub


Sub SetTitles():

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

End Sub



