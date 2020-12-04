Option Explicit

Sub count_referenz()
    Dim i As Integer
    Dim z As Integer
    Dim tmpCell As String
    Dim tmpWeight As Double
    
    Sheets("Sheet1").Select
    
    i = 2
    z = 2
    tmpWeight = 0
    
    Do Until (IsEmpty(Cells(i, 2)) = True)
        tmpCell = Cells(i, 2)
        tmpWeight = Cells(i, 4)
        Do Until (IsEmpty(Cells(z, 2)) = True)
            If tmpCell = Cells(z + 1, 2) Then
                tmpWeight = tmpWeight + Cells(z + 1, 4)
                Rows(z).Delete
                z = z - 1
            End If
            z = z + 1
        Loop
        Cells(i, 4) = tmpWeight
        tmpWeight = 0
        i = i + 1
        z = i
    Loop
    
End Sub
