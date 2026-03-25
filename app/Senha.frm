VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Senha 
   BorderStyle     =   0  'None
   Caption         =   "Seja Bem Vindo!"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   Icon            =   "Senha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Senha.frx":030A
   ScaleHeight     =   3750
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.TextBox txtSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2460
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2520
      Width           =   1635
   End
   Begin VB.ComboBox cboUsuario 
      Height          =   315
      Left            =   2460
      TabIndex        =   0
      Top             =   2040
      Width           =   1635
   End
   Begin MSMask.MaskEdBox mskCPF 
      Height          =   315
      Left            =   2460
      TabIndex        =   1
      Top             =   2040
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   14
      Mask            =   "###.###.###-##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblCopyright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   60
      TabIndex        =   6
      Top             =   3495
      Width           =   270
   End
   Begin VB.Label lblVersao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   4110
      TabIndex        =   5
      Top             =   3495
      Width           =   270
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1575
      TabIndex        =   4
      Top             =   2100
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1740
      TabIndex        =   3
      Top             =   2580
      Width           =   675
   End
End
Attribute VB_Name = "Senha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCfg As ConfigItem
Private moCombo As cComboHelper
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vUltimaData As Date
Dim lNovoCod As Long

Private Sub GerarNovaMensalidade()
sSQL = "SELECT cnpj, razao FROM empresa"
Set r = dbData.OpenRecordset(sSQL)

Dim vCnpj As Integer
Dim vQuantRazao As Integer
If Not r.BOF Then
    vCnpj = SomarDigitos(r("cnpj"))
    vQuantRazao = Len(r("razao"))
End If

'começa a criação
Dim vDataInicio As Date
Dim vDia As Integer
Dim vMes As Integer
Dim vMesInt As String
Dim vAno As Integer
Dim vMesRef As String

'paga proxima data de bloqueio
vDataInicio = Format(DateAdd("m", Val(1), vUltimaData), "dd/mm/yy")

'pega o mês/ano da proxima data de bloquei
vMes = Format(vDataInicio, "m")
vAno = Year(vDataInicio)

'saber o numero do ultimo dia daquele mês
Dim vUltimoDiaMes As Integer
vUltimoDiaMes = Day(DateSerial(vAno, vMes + 1, 0))
vDia = vUltimoDiaMes 'sabe o ultimo dia daquele mês

Dim vDataBloqueio As String
Dim vDataVenc As String

Autonumeracao_Pagamentos

'preenchimento do campo MES/ANO por intenso
vMesInt = Format(vDataInicio, "mmmm")
vAno = Year(vDataInicio)
vMesRef = vMesInt & "/" & vAno

vDataBloqueio = vDia & "/" & vMes & "/" & vAno
vDataVenc = vDia & "/" & vMes & "/" & vAno
vDataBloqueio = Format(DateAdd("d", Val(5), vDataBloqueio), "dd/mm/yy")

'codigo de desbloqueio
    Dim vNumeroMes As Integer
    If vMesInt = "janeiro" Then
        vNumeroMes = 1
    ElseIf vMesInt = "fevereiro" Then
        vNumeroMes = 2
    ElseIf vMesInt = "março" Then
        vNumeroMes = 3
    ElseIf vMesInt = "abril" Then
        vNumeroMes = 4
    ElseIf vMesInt = "maio" Then
        vNumeroMes = 5
    ElseIf vMesInt = "junho" Then
        vNumeroMes = 6
    ElseIf vMesInt = "julho" Then
        vNumeroMes = 7
    ElseIf vMesInt = "agosto" Then
        vNumeroMes = 8
    ElseIf vMesInt = "setembro" Then
        vNumeroMes = 9
    ElseIf vMesInt = "outubro" Then
        vNumeroMes = 10
    ElseIf vMesInt = "novembro" Then
        vNumeroMes = 11
    ElseIf vMesInt = "dezembro" Then
        vNumeroMes = 12
    End If
    
    Dim vCodDesbloqueio As String
    Dim vCodDesbTemp As String
    
    If vNumeroMes Mod 2 = 0 Then
        'MsgBox "Par!"
        vCodDesbloqueio = Left(vCnpj, 1) & "" & Left(vQuantRazao, 1) & "" & Len(vMesInt) & "" & vNumeroMes & "" & UCase(Mid(vMesInt, 3, 1))
    Else
        'MsgBox "Ímpar!"
        vCodDesbloqueio = Mid(vCnpj, 2, 1) & "" & Mid(vQuantRazao, 2, 1) & "" & Len(vMesInt) - 1 & "" & vNumeroMes & "" & UCase(Mid(vMesInt, 2, 1))
    End If

    'Desbloqueio temporario
    If vNumeroMes Mod 2 = 0 Then
        'MsgBox "Par!"
        vCodDesbTemp = Left(vCodDesbloqueio, 1) & "" & Left(vCodDesbloqueio, 1) & "" & vNumeroMes + 1 & "" & UCase(Mid(vMesInt, 4, 1))
    Else
        'MsgBox "Ímpar!"
        vCodDesbTemp = Mid(vCodDesbloqueio, 2, 1) & "" & Mid(vCodDesbloqueio, 2, 1) & "" & Len(vMesInt) - 1 & "" & vNumeroMes + 1 & "" & UCase(Mid(vMesInt, 4, 1))
    End If
    
    dbData.Execute "INSERT INTO  licenca_pagamentos (codigo, dia_vencimento, mes_ref, data_vencimento, data_bloqueio, bloqueio, pago, COD_DESBLOQUEIO, COD_TEMP, Debloqueio_Temp) VALUES (" & _
            lNovoCod & ", " & vDia & ", '" & vMesRef & "', '" & Format$(vDataVenc, "yyyy-dd-MM") & "', '" & Format$(vDataBloqueio, "yyyy-dd-MM") & "', 0, 0, '" & vCodDesbloqueio & "', '" & vCodDesbTemp & "', 0);"
End Sub

Public Function SomarDigitos(CNPJ As String) As Integer
    Dim s As Integer
    Dim i As Integer
    For i = 1 To Len(CNPJ)
      If IsNumeric(Mid(CNPJ, i, 1)) Then
        s = s + Mid(CNPJ, i, 1)
      End If
    Next
    SomarDigitos = s
End Function

'Lista os usuarios permitidos para execução
Sub pListarUsuarios()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT codigo, login FROM usuario WHERE (VISIVEL = 1) ORDER BY login"
   Set r = dbData.OpenRecordset(sSQL)
   
   cboUsuario.Clear
   
   Do While Not r.EOF
      cboUsuario.AddItem r("login")
      cboUsuario.ItemData(cboUsuario.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Entrar()
Static Tag As Byte
Set oCfg = sysConfig("TIPOLOGIN")

Dim codUsuario As Long, nomeUsuario As String
Dim nivelAcesso As Integer
Dim vCodMensalidade As Integer
Dim vDataAtual As Date

If oCfg.Value = "NOME" Then
    If cboUsuario.ListIndex = -1 Then Exit Sub
    If cboUsuario.Text = "" Then Exit Sub
    
    'Recupera o código e o login do usuário
    codUsuario = cboUsuario.ItemData(cboUsuario.ListIndex)
    nomeUsuario = cboUsuario.Text
    
    'Consulta os dados do usuário
    sSQL = "SELECT login, password, nivel FROM usuario WHERE (codigo = " & codUsuario & ")"
    Set r = dbData.OpenRecordset(sSQL)
    
    If r.BOF And r.RecordCount = 0 Then
       ShowMsg "USUÁRIO INEXISTENTE!" & vbCrLf & "O usuário escolhido não existe." & vbCrLf & "Verifique e escolha novamente.", vbInformation
       cboUsuario.ListIndex = -1
       Exit Sub
    End If
    
    If r("password") <> txtSenha Then
       ShowMsg "SENHA ERRADA!" & vbCrLf & "Verifique sua senha e tente novamente", vbInformation
       txtSenha = ""
       Tag = Tag + 1
       
       If Tag = 3 Then
          ShowMsg "Nº DE TENTATIVAS ESGOTADO!" & vbCrLf & "Entre em  contato como o administrador" & vbCrLf & "O sistema será fechado.", vbInformation
          KillProcess "OnlineCommerce"
          Exit Sub
       End If
       
       txtSenha.Text = ""
       txtSenha.SetFocus
       Exit Sub
    End If
    
    'Recupera o nível de acesso do usuário
    nivelAcesso = ValidateNull(r("nivel"))
    
    'Reseta as variáveis
    Tag = 0
    cboUsuario.Text = ""
    txtSenha.Text = ""
    
    'Carrega o form principal
    If nivelAcesso <> 3 Then
        
        'seleciona bloqueio em aberto
        sSQL = "SELECT codigo, bloqueio, mes_ref, data_bloqueio FROM licenca_pagamentos where pago = 0 order by data_bloqueio;"
        Set r = dbData.OpenRecordset(sSQL)

        If Not r.BOF Then
            If r("bloqueio") = 0 Then
                r.MoveFirst
                'se passou da data do bloquei, ele bloqueia
                vCodMensalidade = r("codigo")
                If Date >= r("data_bloqueio") Then
                    dbData.Execute "UPDATE licenca_pagamentos SET bloqueio = 1 WHERE (codigo = " & vCodMensalidade & ");"
                End If
                'Exit Sub
            Else
                'se já estiver bloqueado, pede a senha de desbloqueio
                Unload Me
                Load Senha_Bloqueio
                Senha_Bloqueio.txtMesRef.Text = r("mes_ref")
                Senha_Bloqueio.lblCodMens.Caption = r("codigo")
                Senha_Bloqueio.Show 1
            End If
        Else
        
            sSQL = "SELECT codigo, bloqueio, mes_ref, data_vencimento FROM licenca_pagamentos order by data_vencimento;"
            Set r = dbData.OpenRecordset(sSQL)

            If Not r.BOF Then
                r.MoveLast
                'Dim vDataAtual As Date
                vUltimaData = r("data_vencimento")
                vDataAtual = Date
                If vUltimaData < vDataAtual Then
                    'MsgBox "Falta parcelas"
                    Call GerarNovaMensalidade
                End If
            End If
        End If

        Load Tela_Principal
        vCodFunc = codUsuario
        Tela_Principal.txtCodFuncionario.Text = codUsuario
        Tela_Principal.StatusBar1.Panels(2).Text = nomeUsuario
        Tela_Principal.txtNivel.Text = nivelAcesso
        Unload Senha
        Tela_Principal.Show
    Else
        ShowMsg "USUÁRIO SEM NIVEL DE ACESSO!" & vbCrLf & "Consulte seu nivel de acesso para essa permissão" & vbCrLf & "Consulte o gerente.", vbInformation
        KillProcess "OnlineCommerce"
        Exit Sub
    End If
    
    'Descarrega o form de login
    Unload Me
Else
      sSQL = "SELECT codigo, login, password, nivel FROM usuario WHERE (CPF = '" & mskCPF.Text & "');"
      Set r = dbData.OpenRecordset(sSQL)
      
    codUsuario = ValidateNull(r("codigo"))
    nomeUsuario = ValidateNull(r("login"))
    nivelAcesso = ValidateNull(r("nivel"))
    
      If r.BOF Then
         ShowMsg "USUÁRIO INEXISTENTE!" & vbCrLf & "O usuário escolhido não existe." & vbCrLf & "Verifique e escolha novamente.", vbInformation
         mskCPF.Mask = ""
         mskCPF.Text = ""
         mskCPF.Mask = "###.###.###-##"
         Exit Sub
      End If
      
      If txtSenha = r("password") Then
            If r("nivel") <> 3 Then
                sSQL = "SELECT codigo, bloqueio, mes_ref, data_bloqueio FROM licenca_pagamentos where pago = 0 order by data_bloqueio;"
                Set r = dbData.OpenRecordset(sSQL)
        
                If Not r.BOF Then
                    If r("bloqueio") = 0 Then
                        r.MoveFirst
                        'MsgBox r("data_bloqueio")
                        
                        vCodMensalidade = r("codigo")
                        If Date >= r("data_bloqueio") Then
                            dbData.Execute "UPDATE licenca_pagamentos SET bloqueio = 1 WHERE (codigo = " & vCodMensalidade & ");"
                        End If
                        'Exit Sub
                    Else
                        Unload Me
                        Load Senha_Bloqueio
                        Senha_Bloqueio.txtMesRef.Text = r("mes_ref")
                        Senha_Bloqueio.lblCodMens.Caption = r("codigo")
                        Senha_Bloqueio.Show 1
                    End If
                Else
                
                    sSQL = "SELECT codigo, bloqueio, mes_ref, data_vencimento FROM licenca_pagamentos order by data_vencimento;"
                    Set r = dbData.OpenRecordset(sSQL)
        
                    If Not r.BOF Then
                        r.MoveLast
                        
                        vUltimaData = r("data_vencimento")
                        vDataAtual = Date
                        If vUltimaData < vDataAtual Then
                            'MsgBox "Falta parcelas"
                            Call GerarNovaMensalidade
                        End If
                    End If
                End If
                    
                Load Tela_Principal
                vCodFunc = codUsuario
                Tela_Principal.txtCodFuncionario.Text = codUsuario
                Tela_Principal.StatusBar1.Panels(2).Text = nomeUsuario
                Tela_Principal.txtNivel.Text = nivelAcesso
                Tela_Principal.Show
            Else
                ShowMsg "USUÁRIO SEM NIVEL DE ACESSO!" & vbCrLf & "Consulte seu nivel de acesso para essa permissão" & vbCrLf & "Consulte o gerente.", vbInformation
                KillProcess "OnlineCommerce"
                Exit Sub
            End If
      Else
        Tag = 0
         ShowMsg "SENHA ERRADA!" & vbCrLf & "Verifique sua senha e tente novamente", vbInformation
         txtSenha = ""
         Tag = Tag + 1
         If Tag = 3 Then
            ShowMsg "Nº DE TENTATIVAS ESGOTADO!" & vbCrLf & "Entre em  contato como o administrador" & vbCrLf & "O sistema será fechado.", vbInformation
            End
         End If
         txtSenha.Text = ""
         txtSenha.SetFocus
      End If
    Unload Me
End If

'Fecha a tabela
'If r.State <> 0 Then r.Close
'Set r = Nothing

End Sub
Private Function Autonumeracao_Pagamentos() As Long
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS Ultimo_Pgto FROM licenca_pagamentos;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lNovoCod = r("Ultimo_Pgto") + 1
Else
    lNovoCod = 1
End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Function


Private Sub VerificarPagamentos()
'verificar se tem algum mês bloqueado
sSQL = "SELECT codigo, bloqueio, mes_ref, COD_DESBLOQUEIO FROM licenca_pagamentos;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    If CBool(r("bloqueio")) = True Then
    
       Dim vCodMesBloqueiado As Integer
       Dim vMesRefBloquado As String
       Dim vCodDesbloqueio As String
       ShowMsg "ATENÇÃO!" & vbCrLf & "Entre em contato com o administador do sistema", vbInformation
       vMesRefBloquado = r("mes_ref")
       vCodMesBloqueiado = r("codigo")
       'MsgBox vMesRefBloquado
       vCodDesbloqueio = InputBox("Informe o cód. de desbloqueio do mês: " & vMesRefBloquado & "", "DOWNLOAD XML NFe", "")
    
       If Not Vazio(vCodDesbloqueio) Then
           If vCodDesbloqueio = r("COD_DESBLOQUEIO") Then
               dbData.Execute "UPDATE licenca_pagamentos SET bloqueio = 0, pago = 1, data_liberacao = '" & Format$(Date, "yyyy-dd-MM") & "' WHERE (codigo = " & vCodMesBloqueiado & ");"
           Else
                MsgBox "Código de desbloqueio errado" & vbCrLf & "Entre em contato com o administador do sistema", vbInformation
                Exit Sub
           End If
      End If
       
    
    Else
        MsgBox "desbloqueado"
    End If
End If
 Exit Sub

Dim vDataInicio As Date
Dim vDia As Integer
Dim vMes As Integer
Dim vMesInt As String
Dim vAno As Integer
Dim vMesRef As String

'vDia = txtDia.Text
vMes = Format(Date, "m")
vAno = Year(Date)
Dim vDataBloqueio As String

'Autonumeracao_Pagamentos

'If chkProximo.Value = 1 Then
'    vDataInicio = vDia & " / " & vMes & " / " & vAno
'    vDataInicio = Format(DateAdd("m", Val(1), vDataInicio), "dd/mm/yy")
'    vMesInt = Format(vDataInicio, "mmmm")
'    vAno = Year(vDataInicio)
'    vMesRef = vMesInt & "/" & vAno
'Else
'    vDataInicio = vDia & " / " & vMes & " / " & vAno
'    vMesInt = Format(vDataInicio, "mmmm")
'    vAno = Year(vDataInicio)
'    vMesRef = vMesInt & "/" & vAno
'End If

'vDataBloqueio = Format(DateAdd("d", Val(5), vDataInicio), "dd/mm/yy")

'sSQL = "SELECT codigo FROM licenca_pagamentos;"
'Set r = dbData.OpenRecordset(sSQL)

'If r.BOF Then
'    If r.RecordCount = 0 Then
'        dbData.Execute "INSERT INTO  licenca_pagamentos (codigo, dia_vencimento, mes_ref, data_vencimento, data_bloqueio, bloqueio, pago) VALUES (" & _
'            lNovoCod & ", " & txtDia.Text & ", '" & vMesRef & "', '" & Format$(vDataInicio, "yyyy-dd-MM") & "', '" & Format$(vDataBloqueio, "yyyy-dd-MM") & "', 0, 0);"
'    End If
'End If
End Sub


Private Sub cboUsuario_GotFocus()
   moCombo.AttachTo cboUsuario
End Sub

Private Sub cboUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   KillProcess "OnlineCommerce"
ElseIf KeyCode = 13 Then
   Entrar
End If
End Sub

Private Sub cboUsuario_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboUsuario_Validate(Cancel As Boolean)
Dim sSQL As String
Dim r As ADODB.Recordset
Dim codUsuario As Long

If cboUsuario.ListIndex = -1 Then Exit Sub
If cboUsuario.Text = "" Then Exit Sub

sSQL = "SELECT password FROM usuario WHERE (codigo = " & codUsuario & ");"
Set r = dbData.OpenRecordset(sSQL)

If IsNull(r("password")) Then
   Dim fMod As Senha_Alterar
   Set fMod = New Senha_Alterar
   
   Load fMod
   fMod.CodigoFuncionario = cboUsuario.ItemData(cboUsuario.ListIndex)
   fMod.Show vbModal
   
   Unload fMod
   Set fMod = Nothing
End If
End Sub





Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub

Private Sub mskCPF_GotFocus()
mskCPF.SelStart = 0
mskCPF.SelLength = Len(mskCPF.Text)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   KillProcess "OnlineCommerce"
End If
End Sub

Private Sub Form_Load()
'CenterForm Me, Width, Height
'Set moCombo = New cComboHelper
'pListarUsuarios

lblVersao.Caption = "Versão: " & Trim(str(App.Major)) & "." & Trim(str(App.Minor)) & "." & Trim(str(App.Revision))
lblCopyright.Caption = "Online.Info Sistemas (1998 - " & Year(Now) & ")"
'App.Revision
Dim sSQL As String
Dim r As ADODB.Recordset
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

Set oCfg = sysConfig("TIPOLOGIN")

If App.PrevInstance = True Then
   ShowMsg "O sistema já está em execução nesta máquina!", vbInformation
   End
Else
    If oCfg.Value = "NOME" Then
        cboUsuario.Visible = True
        mskCPF.Visible = False
        Label5.Caption = "Usuário:"
        sSQL = "SELECT codigo, login FROM usuario WHERE (visivel = 1) ORDER BY login;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboUsuario.AddItem r("login")
           cboUsuario.ItemData(cboUsuario.NewIndex) = r("codigo")
           r.MoveNext
        Loop
        
        If r.State <> 0 Then r.Close
        Set r = Nothing
    Else
        cboUsuario.Visible = False
        mskCPF.Visible = True
        Label5.Caption = "CPF:"
        If mskCPF.Visible = True Then mskCPF.SetFocus

    End If
End If


Set oCfg = Nothing
Set moCombo = New cComboHelper
End Sub

Private Sub txtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   KillProcess "OnlineCommerce"
ElseIf KeyCode = 13 Then
   Call Entrar
End If
End Sub
