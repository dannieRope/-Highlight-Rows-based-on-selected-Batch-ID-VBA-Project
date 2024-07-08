Sub HighlightRows()

Dim TotalRow As Integer
Dim RowNumber As Integer

'Getting total number of rows with data excluding the header
   Range("A1").CurrentRegion.Select
   
   TotalRow = Selection.Rows.Count
    
'Loop through the rows that satisfy the conditions and highlight rows green
    For RowNumber = 2 To TotalRow
        If Range("Identifier").Value = Identifier(Selection.Cells(RowNumber, 1).Value) _
            And Key(Selection.Cells(RowNumber, 1).Value) = Range("Key") Then
                Selection.Rows(RowNumber).Interior.ColorIndex = 4
                
                
        End If
        
    Next RowNumber
    
Range("A1").Select

    

End Sub
