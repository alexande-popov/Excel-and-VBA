Attribute VB_Name = "Module1"
Sub Solution1()
Dim i As Long
    ReDim Arr(1 To Range("C11:C20").Rows.Count)
    For i = 1 To Range("C11:C20").Rows.Count
        Arr(i) = Range("C11:C20").Rows(i)
    Next
    [F5].Validation.Delete
    [F5].Validation.Add Type:=xlValidateList, Formula1:=Join(Arr, ",")
End Sub
