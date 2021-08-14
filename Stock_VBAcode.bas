Attribute VB_Name = "Module2"
Sub Stocks()

For Each ws In Worksheets
    
    'Assign Column Titles
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"

    'Create objects to store variables
    Dim ticker As String
    Dim yeardelta As Double
    Dim percentdelta As Double
    Dim volume As Double
    
    Dim stockopen As Double
    Dim stockclose As Double
    
    'Store number of rows as variable
    Dim lastrow As Double
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set volume counter equal to zero before loop
volume = 0

'Set row counter for summary table to 2
Dim sumtab_row As Double
sumtab_row = 2

'Begin for loop starting at row 2 until end of rows
For i = 2 To lastrow

    'If ticker name above does not match current ticker (we are on the FIRST row for this ticker)...
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
         
         'Set ticker name
         ticker = ws.Cells(i, 1).Value
         ws.Range("I" & sumtab_row).Value = ticker
         
         'Add volume to counter
         volume = volume + ws.Cells(i, 7).Value
         
         'Set open price
         stockopen = ws.Cells(i, 3)

    'If ticker name below does not match current ticker (we are on the LAST row for this ticker)...
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then

        'Add volume to counter
        volume = volume + ws.Cells(i, 7).Value

        'Grab stockclose price
        stockclose = ws.Cells(i, 6)
       
       'Need failsafe against dividing by zero...
        If stockopen = 0 Then
            yeardelta = 0
            percentdelta = 0
        Else:
            yeardelta = stockclose - stockopen
            percentdelta = (stockclose - stockopen) / stockopen
        End If

        'Store data in summary table
            ws.Range("L" & sumtab_row).Value = volume
            ws.Range("J" & sumtab_row).Value = yeardelta
            ws.Range("K" & sumtab_row).Value = percentdelta
            ws.Range("K" & sumtab_row).Style = "Percent"
            ws.Range("K" & sumtab_row).NumberFormat = "0.00%"

        'Reset counter for next ticker name
        volume = 0
        sumtab_row = sumtab_row + 1
    
    'If we aren't on a first or last row, just add volume to volume counter
    Else: volume = volume + ws.Cells(i, 7).Value

    End If

Next i

'Change formatting of percent change - negative is red, positive is green
For r = 2 To lastrow

    If ws.Range("J" & r).Value > 0 Then
        ws.Range("J" & r).Interior.ColorIndex = 4

    ElseIf ws.Range("J" & r).Value < 0 Then
        ws.Range("J" & r).Interior.ColorIndex = 3
        
    End If

Next r
    
'BONUS: Add Superlative Summary Table
'Set column titles
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Set row titles
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total volume"

'Create objects to store variables
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double

'Look through column J, Yearly Change, for highest value
For a = 2 To lastrow

    If ws.Cells(a, 11).Value > GreatestIncrease Then
        GreatestIncrease = ws.Cells(a, 11).Value
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("Q2").Style = "Percent"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = ws.Cells(a, 9).Value
    End If

Next a

'Look through column J, Yearly Change, for lowest value
For b = 2 To lastrow
    
    If ws.Cells(b, 11).Value < GreatestDecrease Then
        GreatestDecrease = ws.Cells(b, 11).Value
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("Q3").Style = "Percent"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = ws.Cells(b, 9).Value
    End If
    
Next b

'Look through column L, Total Stock Volume, for highest value
For c = 2 To lastrow
    
    If ws.Cells(c, 12).Value > GreatestVolume Then
        GreatestVolume = ws.Cells(c, 12).Value
        ws.Range("Q4").Value = GreatestVolume
        ws.Range("P4").Value = ws.Cells(c, 9).Value
    End If
  
Next c

'AutoFit columns to make it look nice
ws.Columns("A:Q").AutoFit
    
Next ws

End Sub

