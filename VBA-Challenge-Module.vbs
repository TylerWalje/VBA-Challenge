Attribute VB_Name = "Module1"
Sub stockmarket()

'AddLink()

Dim Open1 As Double
Dim Close1 As Double
Dim Yearly_Change As Double
Dim Ticker_Symbol As String
Dim Percent_Change As Double
Dim Volume As Double
    
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'    Range(1, "I").Value = "Stock"
    Cells(1, "I").Value = "Stock"
    Cells(1, "J").Value = "Yearly_Change"
    Cells(1, "K").Value = "Percent_Change"
    Cells(1, "L").Value = "Total Stock Volume"



Volume = 0

Dim Row As Double

Row = 2

'Dim Column As Long
Dim Column As Integer

Column = 1

Dim i As Long
        
Open1 = Cells(2, Column + 2).Value

For i = 2 To lastrow

    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then

    Ticker_Symbol = Cells(i, Column).Value
    
    Cells(Row, Column + 8).Value = Ticker_Symbol

    Close1 = Cells(i, Column + 5).Value

    Yearly_Change = Close1 - Open1
    
    Cells(Row, Column + 9).Value = Yearly_Change
    
    If (Open1 = 0 And Close1 = 0) Then
        Percent_Change = 0
        
    ElseIf (Open1 = 0 And Close1 <> 0) Then
        Percent_Change = 1
        
    Else
    Percent_Change = Yearly_Change / Open1
    
    Cells(Row, Column + 10).Value = Percent_Change
    
    Cells(Row, Column + 10).NumberFormat = "0.00%"
    End If
    
 
    Volume = Volume + Cells(i, Column + 6).Value
    Cells(Row, Column + 11).Value = Volume
    

    Row = Row + 1

    Open1 = Cells(i + 1, Column + 2)

    Volume = 0

    Else
        Volume = Volume + Cells(i, Column + 6).Value
    End If
    
Next i

        yearlylastrow = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row
        

    For j = 2 To yearlylastrow
    
        If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
        Cells(j, Column + 9).Interior.ColorIndex = 10
        
        ElseIf Cells(j, Column + 9).Value < 0 Then
        Cells(j, Column + 9).Interior.ColorIndex = 3
        
        End If
        
        Next j
        
    Cells(2, Column + 14).Value = "Greatest % Increase"
    
    Cells(3, Column + 14).Value = "Greatest % Decrease"
    
    Cells(4, Column + 14).Value = "Greatest Total Volume"
    
    Cells(1, Column + 15).Value = "Ticker"
    
    Cells(1, Column + 16).Value = "Value"
 
       For m = 2 To yearlylastrow
       
        If Cells(m, Column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & yearlylastrow)) Then
            Cells(2, Column + 15).Value = Cells(m, Column + 8).Value
            Cells(2, Column + 16).Value = Cells(m, Column + 10).Value
            Cells(2, Column + 16).NumberFormat = "0.00%"
            
        ElseIf Cells(m, Column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & yearlylastrow)) Then
            Cells(3, Column + 15).Value = Cells(m, Column + 8).Value
            Cells(3, Column + 16).Value = Cells(m, Column + 10).Value
            Cells(3, Column + 16).NumberFormat = "0.00%"
            
        ElseIf Cells(m, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & yearlylastrow)) Then
            Cells(4, Column + 15).Value = Cells(m, Column + 8).Value
            Cells(4, Column + 16).Value = Cells(m, Column + 11).Value
          
            
            End If
Next m
       Next ws
End Sub


