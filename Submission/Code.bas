Attribute VB_Name = "Module1"
Sub annualReport()

    Dim lastRow As Long
    Dim ticker As String
    Dim yearlyChange As Double
    Dim op As Double
    Dim cp As Double
    Dim prcntChange As Double
    Dim nextRow As Integer
    Dim total As Double
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
'       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Annual Report Table
'       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        'Header titles for annual report
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
    
        ' Initializes opening price row counter
        op = 2
        ' Helper variable to keep track of summary table's row
        nextRow = 2
        ' Total amount charged for each stock type
        total = 0
        ' gets the last row of the data sheet
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To lastRow
            total = total + Cells(i, 7).Value
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                cp = i
                ' Gets ticker
                ticker = Cells(i, 1).Value
                Cells(nextRow, 9).Value = ticker
            
                ' Calcs yearly change
                yearlyChange = Cells(cp, 6).Value - Cells(op, 3).Value
                Cells(nextRow, 10).Value = yearlyChange
                If yearlyChange < 0 Then
                    ' Red
                    Cells(nextRow, 10).Interior.Color = RGB(255, 0, 0)
                ElseIf yearlyChange > 0 Then
                    ' Green
                    Cells(nextRow, 10).Interior.Color = RGB(0, 255, 0)
                End If
            
                ' Calcs percent change of year change
                prcntChange = yearlyChange / Cells(op, 3).Value
                Cells(nextRow, 11).Value = prcntChange
                Cells(nextRow, 11).NumberFormat = "0.00%"
            
                ' Gets total
                Cells(nextRow, 12).Value = total
            
                op = cp + 1
                total = 0
                nextRow = nextRow + 1

                ' Autofit column width
                Columns.AutoFit
            End If
        Next i
    
'       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Summary of Annual Report Table
'       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        Dim lastRowS As Long
        Dim yearlyChangeVal As Range
        Dim totalStkVol As Range
        Dim maxTicker As String
        Dim max As Double
        Dim minTicker As String
        Dim min As Double
        Dim maxTickTot As String
        Dim maxTotal As Double
    
        'Header titles for Summary Table
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Columns("O:O").AutoFit
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
    
        ' Determine the last row in the column
        lastRowS = Cells(Rows.Count, 1).End(xlUp).Row
    
        ' Set the range of the column
        Set yearlyChangeVal = Range("K1:K" & lastRowS)
        Set totalStkVol = Range("L1:L" & lastRowS)
    
        ' Initialize maxVal with the first value in the column
        max = yearlyChangeVal.Cells(2).Value
        min = yearlyChangeVal.Cells(2).Value
        maxTotal = totalStkVol.Cells(2).Value
    
        ' Loop through each row in the column
        For i = 2 To lastRowS
            ' Checks for Largest +/- Percent Change
            If yearlyChangeVal.Cells(i).Value > max Then
                max = yearlyChangeVal.Cells(i).Value
                maxTicker = Cells(i, 9).Value
            ElseIf yearlyChangeVal.Cells(i).Value < min Then
                min = yearlyChangeVal.Cells(i).Value
                minTicker = Cells(i, 9).Value
            End If
            ' Checks for largest volume
            If totalStkVol.Cells(i).Value > maxTotal Then
                maxTotal = totalStkVol.Cells(i).Value
                maxTickTot = Cells(i, 9).Value
            End If
        Next i
    
        ' Fills table
        ' Greatest % Increase
        Cells(2, 16).Value = maxTicker
        Cells(2, 17).Value = max
        Cells(2, 17).NumberFormat = "0.00%"
        'Greatest % Decrease
        Cells(3, 16).Value = minTicker
        Cells(3, 17).Value = min
        Cells(3, 17).NumberFormat = "0.00%"
        'Greatest Total Volume
        Cells(4, 16).Value = maxTickTot
        Cells(4, 17).Value = maxTotal
        Cells(4, 17).NumberFormat = "0.00E+00"
        Columns("P:P").AutoFit
        Columns("Q:Q").AutoFit
    Next ws
End Sub

