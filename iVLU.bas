Function РВПР(SearchValue As Variant, Table As Variant, ResultColumnNum As Long)
    Dim i As Long, res As Long
    Select Case TypeName(Table)
    Case "Range"
        For i = 1 To Table.Rows.Count
            If Table.Cells(i, 1) = SearchValue Then res = i
        Next i
        РВПР = Table.Cells(res, ResultColumnNum)
    Case "Variant()"
        For i = 1 To UBound(Table)
            If Table(i, 1) = SearchValue Then res = i
        Next i
        РВПР = Table(res, ResultColumnNum)
    End Select
End Function

