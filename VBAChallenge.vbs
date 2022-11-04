Attribute VB_Name = "Module1"
Sub VBAChallenge():

    ' Create loop
    For Each ws In Worksheets

        ' Create Column Headers and Value Rows
        ws.Range("L1").Value = "Ticker"
        ws.Range("M1").Value = "Annual Change"
        ws.Range("N1").Value = "Change in Percent"
        ws.Range("O1").Value = "Total Stock Volume"

        ' Define Variables
        Dim Symbol As String
        Dim LastRow As Long
        Dim TotalVolume As Double
        Dim Summary As Long
        Dim StartPrice As Double
        Dim EndPrice As Double
        Dim AnnualChange As Double
        Dim PreviousAmount As Long
        Dim DeltaPercent As Double
        Dim LastRowValue As Long
        
        '    Set variables to baseline
        TotalVolume = 0
        Summary = 2
        PreviousAmount = 2

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
        
            ' Sum the ticker total volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            ' Determine if staying within the same, unique ticker range
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Obtain the ticker symbol and totalvolume and copy to summary table
                Symbol = ws.Cells(i, 1).Value
                ws.Range("L" & Summary).Value = Symbol
                ws.Range("O" & Summary).Value = TotalVolume
                
                ' Return ticker volume total to zero to start over
                TotalVolume = 0

                ' Calculate annual price change by subtracting year end close from year beginning open
                EndPrice = ws.Range("F" & i)
                StartPrice = ws.Range("C" & PreviousAmount)
                AnnualChange = EndPrice - StartPrice
                ws.Range("M" & Summary).Value = AnnualChange

                ' Calculate the delta percent
                If StartPrice = 0 Then
                    DeltaPercent = 0
                Else
                    StartPrice = ws.Range("C" & PreviousAmount)
                    DeltaPercent = AnnualChange / StartPrice
                End If
                
                ' Conditional formatting
                ws.Range("N" & Summary).NumberFormat = "0.00%"
                ws.Range("N" & Summary).Value = DeltaPercent
                If ws.Range("M" & Summary).Value >= 0 Then
                    ws.Range("M" & Summary).Interior.ColorIndex = 4
                Else
                    ws.Range("M" & Summary).Interior.ColorIndex = 3
                End If
            
                ' iterate through summary rows
                Summary = Summary + 1
                PreviousAmount = i + 1
                End If
            Next i
        Next ws
End Sub
