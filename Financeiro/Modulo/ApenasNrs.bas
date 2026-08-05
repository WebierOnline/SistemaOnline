Attribute VB_Name = "mApenasNumeros"
Public Sub ApenasNrs(ByRef Keyasc As Integer)
Select Case Keyasc
    'Se a tecla for numérica (0 - 9) ,backspace (8) ou hifen(-) ou ponto(.)
     Case Asc("0") To Asc("9"), 8, Asc("-"), Asc(".")
        Case Else
        Beep 'Som de erro, nao é necessário
        Keyasc = 0 'Cancela a entrada
End Select
End Sub
