Attribute VB_Name = "mFormataCelular"
Function FormataTelefone(ByVal TEXT As String) As String
Dim i As Long

' ignora vazio
If Len(TEXT) = 0 Then Exit Function

 'verifica valores invalidos
  For i = Len(TEXT) To 1 Step -1
    If InStr("0123456789", Mid$(TEXT, i, 1)) = 0 Then
       TEXT = Left$(TEXT, i - 1) & Mid$(TEXT, i + 1)
    End If
  Next
  ' ajusta a posicao correta
  If Len(TEXT) > 8 And Len(TEXT) < 10 Then
     FormataTelefone = Format$(TEXT, "!@@@@@-@@@@")
  ElseIf Len(TEXT) > 10 And Len(TEXT) < 12 Then
     FormataTelefone = Format$(TEXT, "!(@@) @@@@@-@@@@")
  End If
  
End Function
