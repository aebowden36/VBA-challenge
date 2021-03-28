Sub alphabeticalTesting():

Dim ticker As String
Dim yearchange As Double
Dim percentchange As Double
Dim openCount As Double
Dim closeCount As Double
Dim I As Long

Dim totcount As Double
Dim Summary_Table_Row As Long



For Each ws In Worksheets
    Summary_Table_Row = 2
    Debug.Print (ws.Name)
    ws.Activate
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Value"
    
    
        For I = 2 To LastRow
    
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                ticker = ws.Cells(I, 1).Value
            
                openCount = ws.Cells(2, 3).Value
                closeCount = ws.Cells(I, 6).Value
                yearchange = (closeCount - openCount)
                totcount = totcount + ws.Cells(I, 7).Value
                percentchange = percentchange + (yearchange / openCount)
                    
           
        
            
                ws.Range("I" & Summary_Table_Row).Value = ticker
            
                ws.Range("L" & Summary_Table_Row).Value = totcount
            
                ws.Range("J" & Summary_Table_Row).Value = yearchange
            
                ws.Range("K" & Summary_Table_Row).Value = percentchange
            
                Summary_Table_Row = Summary_Table_Row + 1
            
                totcount = 0
                yearchange = 0
                percentchange = 0
                openCount = Cells(I + 1, 3).Value
                closeCount = Cells(I + 1, 6).Value
                
            
            Else
                
                totcount = totcount + ws.Cells(I, 7).Value
            
            End If
        
    
        Next I
    
    
        For I = 2 To LastRow
        
            ws.Cells(I, "K").Style = "Percent"
                    
        Next I

        For I = 2 To LastRow
        
            If ws.Cells(I, 10).Value > 0 Then
                ws.Cells(I, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(I, 10).Interior.ColorIndex = 3
            End If
            
        
        Next I
    
    
    
Next ws




End Sub
