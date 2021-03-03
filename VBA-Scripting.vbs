Sub Stock_Variables_Sheet()

'NAME COLUMNS
'------------------------------------------

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Stock Volume"

'IDENTIFY VARIABLES
'-----------------------------------------------------

    Dim Stock_Ticker As String
    
    Dim Open_Value As Double
    
    Dim Close_Value As Double

    Dim Yearly_Change As Long
    Yearly_Change = 2
    
    Dim Stock_Volume As Double
    Stock_Volume = 0
    
    Dim Row_Count As Long
    Row_Count = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
'SET VALUES FOR VARIABLES
'-----------------------------------------------------

    For i = 2 To Row_Count

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            Stock_Ticker = Cells(i, 1).Value
            
            Open_Value = Cells(i, 3).Value
            
            Close_Value = Cells(i, 6).Value
        
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
            
'------------------------------------------------------------------
        
            Range("I" & Summary_Table_Row).Value = Stock_Ticker
        
            Range("J" & Summary_Table_Row).Value = Yearly_Value
            
            Range("L" & Summary_Table_Row).Value = Stock_Volume
        
            Yearly_Change = Close_Value - Open_Value
            
            Summary_Table_Row = Summary_Table_Row + 1

'RESET VALUES
'--------------------------------------------------------------
        
            Stock_Value = 0
            
            Yearly_Value = 0
        Else
        
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
            
            Yearly_Value = Close_Value - Open_Value
        End If
    
    Next i
    
'CONDITIONAL FORMATTING
'---------------------------------------------------------------------

    For i = 2 To 3169
        For j = 10 To 10
            If i >= 0 Then
                If j >= 0 Then
                Cells(i, j).Interior.ColorIndex = 3
            Else
                Cells(i, j).Interior.ColorIndex = 4
                End If
            End If
        Next j
    Next i
End Sub


