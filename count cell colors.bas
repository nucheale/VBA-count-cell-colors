Sub colorCount()
Dim colorNames() As String
Dim colorCounts() As Long
Dim n As Integer

Columns("C:E").Delete

lastRow = Cells(1, 1).CurrentRegion.Rows.Count
ReDim Preserve colorNames(0 To 0) As String

n = UBound(colorNames) - 1

For i = 1 To lastRow
    If Not Cells(i, 1).Interior.Color = Cells(i + 1, 1).Interior.Color Then
        contains = False
        For j = 0 To UBound(colorNames)
            If colorNames(j) = Cells(i, 1).Interior.Color Then contains = True
        Next j
        If contains = False Then
            n = n + 1
            ReDim Preserve colorNames(0 To n)
            colorNames(n) = Cells(i, 1).Interior.Color
        End If
    End If
Next i

ReDim colorCounts(0 To n) As Long

For n = 0 To UBound(colorNames)
    Cells(n + 1, 3).Interior.Color = colorNames(n)
Next n

For i = 1 To lastRow
    For n = 0 To UBound(colorNames)
        If Cells(i, 1).Interior.Color = colorNames(n) Then
            colorCounts(n) = colorCounts(n) + 1
        End If
    Next n
Next i

For n = 0 To UBound(colorNames)
    Cells(n + 1, 3).Interior.Color = colorNames(n)
    Cells(n + 1, 3) = colorCounts(n)
'    Cells(n + 1, 4) = "RGB: " & (colorNames(n) Mod 256) & ", " & ((colorNames(n) \ 256) Mod 256) & ", " & colorNames(n) \ 65536
    Cells(n + 1, 4) = (colorNames(n) Mod 256) & ", " & ((colorNames(n) \ 256) Mod 256) & ", " & colorNames(n) \ 65536
    Cells(n + 1, 5) = Hex(colorNames(n))
    Cells(n + 1, 5) = Right(Hex(colorNames(n)), 2) & Right(Left(Hex(colorNames(n)), 4), 2) & Left(Hex(colorNames(n)), 2)
Next n

Columns("A:Z").AutoFit
Columns(3).ColumnWidth = 10

End Sub
