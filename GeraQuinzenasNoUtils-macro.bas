Attribute VB_Name = "M�dulo2"
Sub ConcatenarMesesGerandoQuinzena()
    Dim startRange As Range
    Dim endRange As Range
    Dim resultRange As Range
    Dim startCell As Range
    Dim endCell As Range
    Dim resultCell As Range
    
    ' intervalos
    Set startRange = Range("G2:G25")
    Set endRange = Range("H2:H25")
    Set resultRange = Range("I2:I25")
    
    ' Itera em cada coluna
    For Each startCell In startRange
        
        Set endCell = endRange.Cells(startCell.Row - startRange.Cells(1).Row + 1)
    
        Set resultCell = resultRange.Cells(startCell.Row - startRange.Cells(1).Row + 1)
        
       
        If Month(startCell.Value) = Month(endCell.Value) Then
          
            If Day(startCell.Value) <= 15 Then

                resultCell.Value = "1�Q " & UCase(Left(Format(startCell.Value, "mmm"), 1)) & Mid(Format(startCell.Value, "mmm"), 2)
            Else
        
                resultCell.Value = "2�Q " & UCase(Left(Format(startCell.Value, "mmm"), 1)) & Mid(Format(startCell.Value, "mmm"), 2)
            End If
        Else

            resultCell.Value = ""
        End If
    Next startCell
End Sub




