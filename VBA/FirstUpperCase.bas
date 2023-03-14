Attribute VB_Name = "Module1"
Sub FirstUpperCase()
  Dim strX As String
  For Each x In Selection
    strX = x.Value
    strX = UCase(Left(strX, 1)) & LCase(Mid(strX, 2))
    x.Value = strX
  Next
End Sub
