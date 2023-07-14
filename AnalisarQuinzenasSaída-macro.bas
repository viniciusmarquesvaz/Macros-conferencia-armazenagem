Sub ConverteTabelaQuinzenasBaseSaida()
    Dim dataRange As Range
    Dim cell As Range
    
    
    
 
    Set dataRange = Sheets("BD-Saida").Range("Q2:Q" & Sheets("BD-Saida").Cells(Rows.Count, "Q").End(xlUp).Row)
    'atualiza antes a aba
    Sheets("BD-Saida").ListObjects("Tabela_Consulta_de_DISTRIBUIDORA").QueryTable.Refresh BackgroundQuery:=False
    For Each cell In dataRange
        ' se celula vazia  retorna vazio
        If cell.Value = "" Then
            cell.Value = ""
        Else
            'quinzena correspondente
             If cell.Value >= DateSerial(2022, 11, 1) And cell.Value <= DateSerial(2022, 11, 15) Then
                cell.Value = "1�Q Nov"
            ElseIf cell.Value <= DateSerial(2022, 11, 30) Then
                cell.Value = "2�Q Nov"
            ElseIf cell.Value <= DateSerial(2022, 12, 15) Then
                cell.Value = "1�Q Dez"
            ElseIf cell.Value <= DateSerial(2022, 12, 31) Then
                cell.Value = "2�Q Dez"
            ElseIf cell.Value <= DateSerial(2023, 1, 15) Then
                cell.Value = "1�Q Jan"
            ElseIf cell.Value <= DateSerial(2023, 1, 31) Then
                cell.Value = "2�Q Jan"
            ElseIf cell.Value <= DateSerial(2023, 2, 15) Then
                cell.Value = "1�Q Fev"
            ElseIf cell.Value <= DateSerial(2023, 2, 28) Then
                cell.Value = "2�Q Fev"
            ElseIf cell.Value <= DateSerial(2023, 3, 15) Then
                cell.Value = "1�Q Mar"
            ElseIf cell.Value <= DateSerial(2023, 3, 31) Then
                cell.Value = "2�Q Mar"
            ElseIf cell.Value <= DateSerial(2023, 4, 15) Then
                cell.Value = "1�Q Abr"
            ElseIf cell.Value <= DateSerial(2023, 4, 30) Then
                cell.Value = "2�Q Abr"
            ElseIf cell.Value <= DateSerial(2023, 5, 15) Then
                cell.Value = "1�Q Mai"
            ElseIf cell.Value <= DateSerial(2023, 5, 31) Then
                cell.Value = "2�Q Mai"
            ElseIf cell.Value <= DateSerial(2023, 6, 15) Then
                cell.Value = "1�Q Jun"
            ElseIf cell.Value <= DateSerial(2023, 6, 30) Then
                cell.Value = "2�Q Jun"
            ElseIf cell.Value <= DateSerial(2023, 7, 15) Then
                cell.Value = "1�Q Jul"
            ElseIf cell.Value <= DateSerial(2023, 7, 31) Then
                cell.Value = "2�Q Jul"
            ElseIf cell.Value <= DateSerial(2023, 8, 15) Then
                cell.Value = "1�Q Ago"
            ElseIf cell.Value <= DateSerial(2023, 8, 31) Then
                cell.Value = "2�Q Ago"
            ElseIf cell.Value <= DateSerial(2023, 9, 15) Then
                cell.Value = "1�Q Set"
            ElseIf cell.Value <= DateSerial(2023, 9, 30) Then
                cell.Value = "2�Q Set"
            ElseIf cell.Value <= DateSerial(2023, 10, 15) Then
                cell.Value = "1�Q Out"
             ElseIf cell.Value <= DateSerial(2023, 10, 31) Then
                cell.Value = "2�Q Out"
            ElseIf cell.Value <= DateSerial(2023, 11, 15) Then
                cell.Value = "1�Q Nov"
            ElseIf cell.Value <= DateSerial(2023, 11, 30) Then
                cell.Value = "2�Q Nov"
            ElseIf cell.Value <= DateSerial(2023, 12, 15) Then
                cell.Value = "1�Q Dez"
            ElseIf cell.Value <= DateSerial(2023, 12, 31) Then
                cell.Value = "2�Q Dez"
            End If
        End If
    Next cell
End Sub
