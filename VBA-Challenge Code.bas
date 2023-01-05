Attribute VB_Name = "Module1"
Sub VBA_Challenge()
Dim ws As Worksheet
For Each ws In Worksheets

' Insert headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Yearly Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

' define data parameters
Dim p_o As Double
Dim p_c As Double
Dim Ttl_Change As Double
Dim Pct_Change As Double
Dim Volume As LongLong
Dim row_count As Long

Dim G_increase As Double
Dim G_decrease As Double
Dim G_volume As LongLong
Dim G_increase_T As String
Dim G_decrease_T As String
Dim G_volume_T As String

' locate last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Start values

Volume = 0
row_count = 2
 ' Ttl_Change = p_c - p_o
' Pct_Change = ((p_c - p_o) / p_o)

    G_increase = 0
    G_decrease = 0
    G_volume = 0

' Loop begins

For Row = 2 To LastRow
        
         ' Beginning Operations
        If (ws.Cells(Row, 1).Value <> ws.Cells(Row - 1, 1)) Then
       p_o = ws.Cells(Row, 3)
       End If
       
        
        ' Every Time Operation
        
        Volume = Volume + ws.Cells(Row, 7).Value

        ' End Operations
       If (ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1)) Then
       p_c = ws.Cells(Row, 6).Value
       
       
Ttl_Change = p_c - p_o
Pct_Change = ((p_c - p_o) / p_o)

       ' transcribe values
       ws.Cells(row_count, 9).Value = ws.Cells(Row, 1).Value
       ws.Cells(row_count, 10).Value = Ttl_Change
       ws.Cells(row_count, 10).NumberFormat = "0.00"
       ws.Cells(row_count, 11).Value = Pct_Change
        ws.Cells(row_count, 11).NumberFormat = "0.00%"
       ws.Cells(row_count, 12).Value = Volume
       
       ' tag output value by color
       If Ttl_Change > 0 Then
       ws.Cells(row_count, 10).Interior.ColorIndex = 4
       Else: ws.Cells(row_count, 10).Interior.ColorIndex = 3
       End If
       
        ' update g_increase and g_decrease
       If Pct_Change > G_increase Then
       G_increase = Pct_Change
       G_increase_T = ws.Cells(row_count, 9).Value
       ElseIf Pct_Change < G_decrease Then
        G_decrease = Pct_Change
        G_decrease_T = ws.Cells(row_count, 9).Value
        End If
        
      ' update g_volume
       If Volume > G_volume Then
       G_volume = Volume
       G_volume_T = ws.Cells(row_count, 9).Value
        End If
        
       ' reset and prepare for new values
       Volume = 0
        row_count = row_count + 1
        
       End If
       
       ' End Loop
Next Row

' output 'greatest' values
ws.Range("Q2").Value = G_increase
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("P2").Value = G_increase_T
ws.Range("Q3").Value = G_decrease
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("P3").Value = G_decrease_T
ws.Range("Q4").Value = G_volume
ws.Range("P4").Value = G_volume_T
    
Next ws

End Sub


