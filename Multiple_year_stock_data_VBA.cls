VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_analysis():
    'define variables
    
    Dim total As LongLong
    Dim row As Long
    Dim ws As Worksheet
    Dim opening_price As Double
    Dim closing_price As Double
    Dim rowCount As Long
    Dim outCounter As Long
    
    Dim change_high As Double
    Dim change_low As Double
    Dim total_greatest As LongLong
    Dim ticker_high As String
    Dim ticker_low As String
    Dim ticker_vol As String
    
    Dim debug_counter As LongLong
    
            
    
    For Each ws In Worksheets
        
        'Set labels
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
        outCounter = 2
        For row = 2 To rowCount
        'outCounter = outCounter + 1
            
            'when current row is NEW ticker
            If ws.Cells(row, 1).Value <> ws.Cells(row - 1, 1).Value Then
                'collect initial values and print to new cell
                ws.Cells(outCounter, 9).Value = ws.Cells(row, 1).Value 'ticker
                opening_price = ws.Cells(row, 3).Value 'opening price for the year
                total = ws.Cells(row, 7) 'initial volume
                outCounter = outCounter + 1
                                 
            'when current row is LAST of current ticker
            ElseIf ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then
                'MsgBox ("Else triggered")
                closing_price = ws.Cells(row, 6).Value
                'MsgBox (closing_price)
                ws.Cells(outCounter - 1, 10).Value = closing_price - opening_price
                'MsgBox (3)
                ws.Cells(outCounter - 1, 11).Value = FormatPercent((ws.Cells(outCounter - 1, 10).Value / opening_price), 2)
                'MsgBox (4)
                total = total + ws.Cells(row, 7)
                'MsgBox (5)
                ws.Cells(outCounter - 1, 12).Value = total
                'MsgBox (6)
                                
            'anything inbetween
            Else
                'MsgBox ("Else triggered")
                total = total + ws.Cells(row, 7)
            
            End If
        
        Next row
        
        change_high = 0
        change_low = 0
        total_greatest = 0
        
        rowCount = ws.Cells(Rows.Count, "I").End(xlUp).row
        For row = 2 To rowCount
                                           
            If ws.Cells(row, 11).Value > change_high Then
                change_high = ws.Cells(row, 11).Value
                ticker_high = ws.Cells(row, 9).Value
                
            End If
                
            If ws.Cells(row, 11).Value < change_low Then
                change_low = ws.Cells(row, 11).Value
                ticker_low = ws.Cells(row, 9).Value
            End If
            
            If ws.Cells(row, 12).Value > total_greatest Then
                total_greatest = ws.Cells(row, 12).Value
                ticker_vol = ws.Cells(row, 9).Value
            End If
                               
            'Format Yearly Change
            If ws.Cells(row, 10).Value > 0 Then
                ws.Cells(row, 10).Interior.ColorIndex = 4
            
            End If
            
            If ws.Cells(row, 10).Value < 0 Then
                ws.Cells(row, 10).Interior.ColorIndex = 3
            
            End If
            
            'Format Percent Change
            If ws.Cells(row, 11).Value > 0 Then
                ws.Cells(row, 11).Interior.ColorIndex = 4
            
            End If
            
            If ws.Cells(row, 11).Value < 0 Then
                ws.Cells(row, 11).Interior.ColorIndex = 3
            
            End If
            
                
        Next row
                
        ws.Cells(2, 16).Value = ticker_high
        ws.Cells(2, 17).Value = FormatPercent(change_high, 2)
        ws.Cells(3, 16).Value = ticker_low
        ws.Cells(3, 17).Value = FormatPercent(change_low, 2)
        ws.Cells(4, 16).Value = ticker_vol
        ws.Cells(4, 17).Value = total_greatest
        
    Next ws
    
End Sub


