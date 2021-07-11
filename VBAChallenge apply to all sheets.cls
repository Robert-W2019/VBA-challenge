VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_analysis_allsheets():

    'set dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim ws As Worksheet
    
    
    For Each ws In Worksheets
        'set values for each worksheet
        j = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0
        
        
        'set title rows
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'get the row number of the last row of data
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To rowCount
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'store results
                total = total + ws.Cells(i, 7).Value
                
                If total = 0 Then
                    ws.Range("i" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("j" & 2 + j).Value = 0
                    ws.Range("k" & 2 + j).Value = "%" & 0
                    ws.Range("l" & 2 + j).Value = 0
                    
                Else
                    'find first non zero starting value
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                    
                    'Calc change
                    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percentChange = Round((change / ws.Cells(start, 3) * 100), 2)
                    
                    'start the next start ticker
                    start = i + 1
                    
                    'print the results to a& 2+j).value = ws.cells(i,1).value
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = Round(change, 2)
                    ws.Range("K" & 2 + j).Value = "%" & percentChange
                    ws.Range("L" & 2 + j).Value = total
                    
                    
                    If change > 0 Then
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    
                    
                    ElseIf change < 0 Then
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        
                    Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                        
                
                    
                    End If
                    
                 End If
                 
                    
                'reset variables for new stocks
                total = 0
                change = 0
                j = j + 1
                days = 0
                dailyChange = 0
                
            Else
                total = total + ws.Cells(i, 7).Value
            End If
            
        Next i
            

        
    Next ws
        
    







End Sub
