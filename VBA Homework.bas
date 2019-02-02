Attribute VB_Name = "Module3"
Sub Stock_data()

 Dim Stock_data As Integer
 Dim ws As Worksheets


 Dim ticker As String
 Dim Total As Double
 Dim volume As Double
 Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
'Label Header
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    
   
   For i = 2 To RowCount
   RowCount = Cells(Rows.Count, "a").End(x1Up).Row
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ticker = ws.Cells(i, 1).Value
            
            Total = Total + Cells(i, 7).Value
            
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            
            Range("J" & 2 + j).Value = Total
            
            
            Total = 0
            ticker = 0
            
            j = j + 1
            i = i + 1
            
            
        Else
         Total = Total + Cells(i, 7).Value
         
        End If
        
    Next i
End Sub

    
    
