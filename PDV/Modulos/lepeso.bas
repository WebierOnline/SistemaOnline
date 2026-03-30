Attribute VB_Name = "lepeso"
Global portaCOM As String
Global vDigitoInicial As Integer
Global vQuantDigitos As Integer
Public Sub abrePortaSerial(canalserial As String)
If PDV.MSComm1.PortOpen Then
    PDV.MSComm1.PortOpen = False
End If
 
On Error GoTo fim

portaCOM = canalserial

PDV.MSComm1.CommPort = canalserial
PDV.MSComm1.Settings = "9600,n,8,2"
PDV.MSComm1.InputLen = 0
PDV.MSComm1.PortOpen = True
Exit Sub
fim:
  teste = MsgBox("Porta Serial COM" & canalserial & " não pode ser aberta", vbCritical, "ERRO")
End Sub
Public Sub fechaPortaSerial()
'If PDV.MSComm1.PortOpen Then
 PDV.MSComm1.PortOpen = False
'End If
End Sub
Public Function enviaComandoSerial(comando As String)
If (PDV.MSComm1.PortOpen) Then
    PDV.MSComm1.Output = comando
    msgSendToECF (comando)
Else
    teste = MsgBox("Porta Serial Fechada!" & Chr(13) & "Configure a porta serial antes de enviar os comandos.", vbCritical, "ERRO")
End If
End Function

Public Sub Delay(t As Integer)
PDV.Timer2.Interval = t
PDV.Timer2.Enabled = True
Do While PDV.Timer2.Enabled
    DoEvents
Loop
End Sub
Public Sub msgSendToECF(comando As String)
'PDV.Text2.Text = Val(comando)
PDV.Text2.Text = comando
End Sub
Public Sub returnOfECF(comando As String)
Dim vPesoCerto As Double

'vPesoCerto = Mid(comando, 2, 5)   'desativei no dia que estava configurando a balança Urano
vPesoCerto = Mid(comando, vDigitoInicial, vQuantDigitos)
vPesoCerto = Val(vPesoCerto)
vPesoCerto = vPesoCerto / 1000
'PDV.Text1.Text = Val(comando)  'desativei no dia que estava configurando a balança Urano
'PDV.Text1.Text = comando   'desativei no dia que estava configurando a balança Urano
'PDV.txtQuant.Text = Mid(comando, 19, 6)  'desativei no dia que estava configurando a balança Urano
PDV.txtQuant.Text = FormatNumber(vPesoCerto, 3)
'PDV.txtQuant.Text = comando  'desativei no dia que estava configurando a balança Urano
If vPedirPeso = True Then SendKeys "{ENTER}"
End Sub
