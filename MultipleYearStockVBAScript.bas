Attribute VB_Name = "Module1"
'Create a script that will loop through all the stocks for one year and output the following information.
Sub VBAChallenge()

    Dim Ticker As String
    Ticker = " "
    Dim YearlyChange As Double
    Dim Opening As Double
    Opening = Cells(2, 3).Value
    Dim Closing As Double
    Closing = 0
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
 
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    

    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
For i = 2 To LastRow
  

    
'The ticker symbol.
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker = Cells(i, 1).Value
    
        Range("I" & Summary_Table_Row).Value = Ticker
 
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
        Closing = Cells(i, 6).Value
        YearlyChange = Closing - Opening
        Range("J" & Summary_Table_Row).Value = YearlyChange
'Change color of YearlyChange based on positive or negative change
        If (YearlyChange > 0) Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
            
        
        
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        If Opening <> 0 Then
        PercentChange = (Closing / Opening) - 1
        End If
        Range("K" & Summary_Table_Row).Value = PercentChange
        
 'Found this "Opening" bit of code here https://freesoft.dev/program/163047389 - Only referenced this site when noted in the code
        Opening = Cells(i + 1, 3).Value
'The total stock volume of the stock.
       TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
       
       Range("L" & Summary_Table_Row).Value = TotalStockVolume


    Summary_Table_Row = Summary_Table_Row + 1
    TotalStockVolume = 0
    YearlyChange = 0
    PercentChange = 0
End If

    TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
    

Next i




End Sub





