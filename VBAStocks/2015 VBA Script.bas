Attribute VB_Name = "Module111"
Sub YearlyStockData()
    Dim openingValue As Long
    Dim closingValue As Long
    Dim OpeningDate As Long
    Dim ClosingDate As Long
    Dim volume As Long
    Dim table_row As Double
    Dim yearly_change As Double
    Dim percent_change As Long
    
    
    
    'find the last row
    'LastRow = ws.A(Rows.Count, 1).End(xlUp).Row
    
    'loop through table
    table_row = 2
    
    'find the total volume
    volume = 0
 
    
    'create OpeningDate and ClosingDate
        OpeningDate = 20150101
        ClosingDate = 20151230
        
   'loop through the ticker column
    For i = 2 To 70926
    
    'list the 1rst ticker name
    Cells(2, 9).Value = Cells(2, 1).Value
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
       
        'list the new ticker in the next empty space
         ticker = Cells(i, 1).Value
        Range("I" & table_row).Value = ticker
        
        'add to the total volume of stock
        volume = volume + Cells(i, 7).Value
        
        'print the total volume
        Range("L" & table_row).Value = volume
        
        'reset volume to 0
        volume = 0
        
        'add one to the table_row
        table_row = table_row + 1
        
    End If
    
    'find and store the openingValue
    If Cells(i, 2) = OpeningDate Then
       
        openingValue = Cells(i, 3)
       
    End If
       
    'find and store the closingValue
    If Cells(i, 2) = ClosingDate Then
       
        closingValue = Cells(i, 6)
       
    End If
       
    yearly_change = (openingValue) - (closingValue)
    Range("J" & table_row).Value = yearly_change
    
    
    'calculate percent change
        percent_change = ((Range("J" & table_row).Value) / (openingValue)) * 100
        Range("K" & table_row).Value = percent_change
        
        If openingValue = 0 Then
        
        percent_change = closingValue
        
        End If
        
    'color the negative and positive changes
    If Range("K" & table_row).Value >= 0 Then
        Range("K" & table_row).Interior.ColorIndex = 4
    
    Else
        Range("K" & table_row).Interior.ColorIndex = 3
    End If
    
        
       
    
    
            
    
    Next i
    
    
    
    
End Sub
