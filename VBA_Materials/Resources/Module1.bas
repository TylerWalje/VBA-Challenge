Attribute VB_Name = "Module1"
Sub analyzeStocksAndCreateSummary()
    Dim ws As Worksheet
    Dim lastrow As Long, k As Long, summaryrow As Long
    Dim startprice As Double, endprice As Double
    Dim yearlychange As Double, percentchange As Double, totalvolume As Double
    Dim ticker As String
    Dim maxincrease As Double, maxdecrease As Double, maxvolume As Double
    Dim maxincreaseticker As String, maxdecreaseticker As String, maxvolumeticker As String
    
    maxincrease = 0
    maxdecrease = 0
    maxvolume = 0

    ' Loop through worksheets
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "2018" Or ws.Name = "2019" Or ws.Name = "2020" Then
            
            lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            summaryrow = 2
            totalvolume = 0
            startprice = 0
            
            ' Add headers
            If ws.Cells(1, 9).Value = "" Then
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Yearly Change"
                ws.Cells(1, 11).Value = "Percent Change"
                ws.Cells(1, 12).Value = "Total Stock Volume"
            End If

            ' Loop through all rows
            For k = 2 To lastrow

                totalvolume = totalvolume + ws.Cells(k, 7).Value


                If k = lastrow Or ws.Cells(k + 1, 1).Value <> ws.Cells(k, 1).Value Then
                    endprice = ws.Cells(k, 6).Value
                    
                    ' Yeary change
                    yearlychange = endprice - startprice
                    
                    ' Percent change
                    If startprice <> 0 Then
                        percentchange = yearlychange / startprice
                    Else
                        percentchange = 0
                    End If
                    
                    ' New columns
                    ws.Cells(summaryrow, 9).Value = ticker
                    ws.Cells(summaryrow, 10).Value = yearlychange
                    ws.Cells(summaryrow, 11).Value = percentchange
                    ws.Cells(summaryrow, 12).Value = totalvolume
                    
                    ' Format
                    ws.Cells(summaryrow, 11).NumberFormat = "0.00%"
                    
                    If percentchange > maxincrease Then
                        maxincrease = percentchange
                        maxincreaseticker = ticker
                    ElseIf percentchange < maxdecrease Then
                        maxdecrease = percentchange
                        maxdecreaseticker = ticker
                    End If
                    If totalvolume > maxvolume Then
                        maxvolume = totalvolume
                        maxvolumeticker = ticker
                    End If
                    
                    summaryrow = summaryrow + 1
                    If k < lastrow Then
                        startprice = ws.Cells(k + 1, 3).Value
                        ticker = ws.Cells(k + 1, 1).Value
                    End If
                    totalvolume = 0
                ElseIf k = 2 Then
                    ' starting price
                    startprice = ws.Cells(k, 3).Value
                    ticker = ws.Cells(k, 1).Value
                End If
            Next k
            
            ' Conditional formatting
            Dim rangeyearlychange As Range
            Dim rangepercentchange As Range

            Set rangeyearlychange = ws.Range("J2:J" & summaryrow - 1)
            Set rangepercentchange = ws.Range("K2:K" & summaryrow - 1)

            ' Run code without earlier formatting
            rangeyearlychange.FormatConditions.Delete
            rangepercentchange.FormatConditions.Delete

            ' Yeary change
            With rangeyearlychange
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
                .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 255, 0)
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
                .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)
            End With

            ' Percent change
            With rangepercentchange
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
                .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 255, 0)
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
                .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)
            End With
            
        End If
    Next ws

    ' New worksheet
    Dim summarysheet As Worksheet
    Set summarysheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    summarysheet.Name = "Summary"

    With summarysheet
    
        ' Headers
        .Cells(1, 1).Value = "Category"
        .Cells(1, 2).Value = "Ticker"
        .Cells(1, 3).Value = "Value"

        ' Values
        .Cells(2, 1).Value = "Greatest % Increase"
        .Cells(2, 2).Value = maxincreaseticker
        .Cells(2, 3).Value = maxincrease
        .Cells(2, 3).NumberFormat = "0.00%"

        .Cells(3, 1).Value = "Greatest % Decrease"
        .Cells(3, 2).Value = maxdecreaseticker
        .Cells(3, 3).Value = maxdecrease
        .Cells(3, 3).NumberFormat = "0.00%"

        ' Format
        .Cells(4, 1).Value = "Greatest Total Volume"
        .Cells(4, 2).Value = maxvolumeticker
        .Cells(4, 3).Value = maxvolume
        .Cells(4, 3).NumberFormat = "#,##0"
    End With
End Sub

