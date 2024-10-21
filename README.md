Private Function RangeNameExists(ActiveWorkbook As Excel.Workbook, nname As String) As Boolean
        Dim n As Excel.Name
        RangeNameExists = False
        For Each n In ActiveWorkbook.Names
            If n.Name = nname Then
                RangeNameExists = True
                Exit Function
            End If
        Next n
    End Function
