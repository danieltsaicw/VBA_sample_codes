Sub Sort2()
     
    Application.Calculation = xlManual
    MsgBox ("Sorting ...")
    Worksheets("Live List").Activate
    
    lastRowCurrent = Worksheets("Live List").Range("I4").End(xlDown).Row
    Range("5:" & lastRowCurrent).Sort key1:=Range("O5:O" & lastRowCurrent), Order1:=xlAscending
    
           
    ' sort all by EV/EBIT
    Dim cell As Range
    Do While Range("O5") = 0
        
        Set cell = Range("O5")
        cell.Clear
        
        If Range("O" & cell.Row) > 0 Then
            'cell.Value = iRow
            'iRow = iRow + 1
        Else
            'Range("I4").End(xlDown).Row
            
            cell.EntireRow.Select
            Selection.Cut
            'cell.EntireRow.Cut
            Range("A" & Range("I4").End(xlDown).Row + 1).EntireRow.Insert
        
        End If
        cell.NumberFormat = "General"
    Loop
    
        
    ' re-Sort top 100 by Delta
    Range("5:" & 5 + 100 - 1).Sort key1:=Range("V5:V" & 5 + 100 - 1), Order1:=xlAscending
    
    
    
    ' mark rank number
    iRow = 1
    For Each cell In Range("G5:G" & Range("G5").End(xlDown).Row)
        cell.Clear
        If Range("O" & cell.Row) > 0 Then
            cell.Value = iRow
            iRow = iRow + 1
        End If
        cell.NumberFormat = "General"
    Next cell
    
    
    
    Columns("G").AutoFit
    
    
    
    MsgBox ("Sort done")
End Sub
