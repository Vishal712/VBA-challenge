Attribute VB_Name = "Module11"
Sub VBAChallengewithEC()
    Dim i As Long
    Dim ticker As String
    'tickercount counts how many rows to go for the output'
    Dim tickercount As Long
    tickercount = 2
    Dim totalvol As Double
    totalvol = 0
    Dim opening As Double
    opening = -999 'this stores the opening stock
    Dim closing As Double
    closing = 0 'this stores the closing stock
    Dim percentage As Double
    Dim change As Double
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    'EC variables for output'
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    'Setting these values so they start with the first value found'
    GreatestIncrease = -999
    GreatestDecrease = 1
    GreatestVolume = -1
For Each ws In Worksheets
    'Creating the labels for new output'
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Char"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    'checking if next row is the same'
    For i = 2 To LastRow
        If (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) Then
            ticker = ws.Cells(i, 1).Value
            totalvol = totalvol + ws.Cells(i, 7).Value
            If (opening = -999) Then
                opening = ws.Cells(i, 3).Value
            End If
            
        'now checking if it is different, write to ouput and reset values'
        Else
            totalvol = totalvol + ws.Cells(i, 7).Value
            ws.Cells(tickercount, 9).Value = ticker
            ws.Cells(tickercount, 12).Value = totalvol
            closing = ws.Cells(i, 6).Value
            change = closing - opening
            ws.Cells(tickercount, 10).Value = change
            If (change = 0) Then
                percentage = 0
            Else
                If (opening = 0) Then
                    percentage = 0
                Else
                    percentage = change / opening
                End If
                
            End If
            
            ws.Cells(tickercount, 11).Value = percentage
            ws.Cells(tickercount, 11).NumberFormat = "0.00%"
            If (change > 0) Then
                ws.Cells(tickercount, 10).Interior.ColorIndex = 4
            End If
            If (change < 0) Then
                ws.Cells(tickercount, 10).Interior.ColorIndex = 3
            End If
            
            'Checking for Greatest Values'
            If (totalvol > GreatestVolume) Then
                GreatestVolume = totalvol
                GreatestVolumeTicker = ticker
            End If
            If (percentage > 0) Then
                If (percentage > GreatestIncrease) Then
                    GreatestIncrease = percentage
                    GreatestIncreaseTicker = ticker
                End If
            End If
            If (percentage < 0) Then
                If (percentage < GreatestDecrease) Then
                    GreatestDecrease = percentage
                    GreatestDecreaseTicker = ticker
                End If
            End If
           
            
            tickercount = tickercount + 1
            totalvol = 0
            opening = -999
        End If
    Next i
    
    'Creating Cells for Total Values EC'
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = GreatestIncreaseTicker
    ws.Cells(2, 17).Value = GreatestIncrease
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = GreatestDecreaseTicker
    ws.Cells(3, 17).Value = GreatestDecrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = GreatestVolumeTicker
    ws.Cells(4, 17).Value = GreatestVolume

'Reseting Variables for the next worksheet so there is no overlap'
GreatestIncrease = -999
GreatestDecrease = 1
GreatestVolume = -1
tickercount = 2
change = 0
opening = -999
Next ws


End Sub
