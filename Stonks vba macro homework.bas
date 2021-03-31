Attribute VB_Name = "Module1"
Sub Stonks():

    For Each ws In Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        Dim ticker As String
        Dim lastrow As Long
        Dim totaltickervolume As Double
        totaltickervolume = 0
        Dim summarytablerow As Long
        summarytablerow = 2
        Dim yearlyopen As Double
        Dim yearlyclose As Double
        Dim yearlychange As Double
        Dim previousamount As Long
        previousamount = 2
        Dim percentchange As Double
        Dim lastrowvalue As Long

        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow

            totaltickervolume ws.Cells(i, 7).Value = totaltickervolume

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                tickerName = ws.Cells(i, 1).Value
                ws.Range("I" & summarytablerow).Value = tickerName
                ws.Range("L" & summarytablerow).Value = totaltickervolume
                totaltickervolume = 0

                yearlyopen = ws.Range("C" & previousamount)
                yearlyclose = ws.Range("F" & i)
                yearlychange = yearlyclose - yearlyopen
                ws.Range("J" & summarytablerow).Value = yearlychange

                If yearlyopen = 0 Then
                    percentchange = 0
                Else
                    yearlyopen = ws.Range("C" & previousamount)
                    percentchange = yearlychange / yearlyopen
                End If

                ws.Range("K" & summarytablerow).NumberFormat = "0.00%"
                ws.Range("K" & summarytablerow).Value = percentchange

                If ws.Range("J" & summarytablerow).Value >= 0 Then
                    ws.Range("J" & summarytablerow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & summarytablerow).Interior.ColorIndex = 3
                End If
            
                summarytablerow = summarytablerow + 1
                previousamount = i + 1
                End If
            Next i

    Next ws

End Sub
