Attribute VB_Name = "Module1"
Sub stock_analyzer()
For Each ws In Worksheets

    ws.Activate
    
    'define the headers for each field
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Sotck Volume"
    
    
    Dim Ticker As Double
    Dim Ticker_Name As String
    Dim Summary_Table_Row As Double
    Dim Volume As Double
    Dim Yearly_Open As Double
    Dim Yearly_Close As Double
    Dim Yearly_Change As Double
    Dim LastRow As Double
    
    
    Yearly_Open = Cells(2, 3).Value
    
    Volume = 0
    
    Summary_Table_Row = 2
    
    'Determine the Last Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        'Print the last row of the sheet
        'MsgBox (LastRow)
    
   'loop throug all the ticker names
    For i = 2 To LastRow
    
        Current_Ticker = ws.Cells(i, 1).Value
        Next_Ticker = ws.Cells(i + 1, 1).Value
              
        'add the ticker total volume
        Volume = Volume + ws.Cells(i, 7).Value
            
        'check if we are still within the same ticker symble
        If Current_Ticker <> Next_Ticker Then
            
            'Check if we are still within the same ticker name
            Ticker_Name = Current_Ticker
                'Debug.Print (Current_Ticker)
                    
            'Print the tickername in the summry table
            Range("I" & Summary_Table_Row).Value = Ticker_Name
    
            'print the total ticker volume in the summaary table row
            Range("L" & Summary_Table_Row).Value = Volume
                
                'Debug.Print ("Total " + Str(Volume))
                                 
            'reset ticker total
            Volume = 0
     
            Yearly_Close = Range("F" & i)
            Yearly_Change = Yearly_Close - Yearly_Open
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            'Conditional formatting that will highlight positive change in green and negative change in red
            ColorRed = 3
            ColorGreen = 4

            If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                 ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = ColorGreen
                         
            ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                 ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = ColorRed
                 
            End If
              
            If Yearly_Open = 0 Then
                    Percent_Change = 0
            Else
                    Percent_Change = (Yearly_Close - Yearly_Open) / Yearly_Open
            End If
            
            
            'reset the yealy open to the next ticker
            Yearly_Open = Cells(i + 1, 3).Value
            
            'printing the value of percent change and cell formatting
            Range("K" & Summary_Table_Row).Value = Percent_Change
            Range("K" & Summary_Table_Row).Style = "Percent"
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            
             'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
             'Debug.Print ("Summary Row Number " + Str(Summary_Table_Row))
            
        End If
       
   Next i
        
    Next ws

End Sub



