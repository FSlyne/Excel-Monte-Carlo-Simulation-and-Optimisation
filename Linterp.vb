Public Function Linterp(Tbl As Range, x As Double) As Variant 
     ' linear interpolator / extrapolator
     ' Tbl is a two-column range containing known x, known y, sorted x ascending
     
    Dim nRow As Long 
    Dim iLo As Long, iHi As Long 
     
    nRow = Tbl.Rows.Count 
    If nRow < 2 Or Tbl.Columns.Count <> 2 Then 
        Linterp = CVErr(xlErrValue) 
        Exit Function '-------------------------------------------------------->
    End If 
     
    If x < Tbl(1, 1) Then ' x < xmin, extrapolate from first two entries
        iLo = 1 
        iHi = 2 
    ElseIf x > Tbl(nRow, 1) Then ' x > xmax, extrapolate from last two entries
        iLo = nRow - 1 
        iHi = nRow 
    Else 
        iLo = Application.Match(x, Application.Index(Tbl, 0, 1), 1) 
        If Tbl(iLo, 1) = x Then ' x is exact from table
            Linterp = Tbl(iLo, 2) 
            Exit Function '---------------------------------------------------->
        Else ' x is between tabulated values, interpolate
            iHi = iLo + 1 
        End If 
    End If 
     
    Linterp = Tbl(iLo, 2) + (Tbl(iHi, 2) - Tbl(iLo, 2)) _ 
    * (x - Tbl(iLo, 1)) / (Tbl(iHi, 1) - Tbl(iLo, 1)) 
     
End Function 
