Sub VBA_StockData():

    ' Create Dimensions needed throughout code
    Dim ws As Worksheet
    Dim total_volume As Double
    Dim i As Integer
    Dim j As Integer

    ' Make the script loop through each worksheet of the workbook
    For Each ws In Worksheets

        ' Re-set each variable to 0 when the script re-runs for a different worksheet
        total_volume = 0
        j = 0

        ' Name columns where the new data will populate
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"

        ' Make each column title bold
        ws.Range("I1").Font.Bold = True
        ws.Range("J1").Font.Bold = True

        ' Gain the last row available with data in order to use for range during for loop
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To LastRow

            ' If the ticker in column A changes, then do the following:
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Add in the last volume amount of that ticker to the total volume 
                total_volume = total_volume + ws.Cells(i, 7).Value

                ' Display the ticker under the Ticker column and have each new ticker print in the cell beneath by having j+2
                ws.Range("I" & j + 2).Value = ws.Cells(i, 1).Value

                ' Display the total volume under the Total Stock Volume column and have each total volume print in the cell beneath by having j+2
                    ' The total volume will print next to the corresponding ticker
                ws.Range("J" & j + 2).Value = total_volume

                ' Reset the total volume to 0 to start calculating new volume for new ticker
                total_volume = 0

                ' Move on to the next row in order to display data correctly in column
                j = j + 1


            ' Else (if the ticker is still the same)
            Else

                ' Keep adding up the volumes in order to gain total 
                total_volume = total_volume + ws.Cells(i, 7).Value

            End If

        ' Move on to the next row within range of the For Loop
        Next i

    ' Move to the next worksheet in workbook 
    Next ws

End Sub
