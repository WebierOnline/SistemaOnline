VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Caixa_Fechamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SITUAÇÃO DO CAIXA"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "Caixa_Fechamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   5955
      Begin VB.TextBox txtCodCaixa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   420
         Width           =   1875
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3270
         TabIndex        =   20
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. do Caixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   180
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1035
      Left            =   60
      TabIndex        =   8
      Top             =   1920
      Width           =   5955
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   11
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saída"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   10
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lblSaida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label lblEntrada 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   1875
      End
   End
   Begin VB.Frame Frame7 
      Height          =   915
      Left            =   60
      TabIndex        =   7
      Top             =   960
      Width           =   5955
      Begin VB.TextBox txtCodFuncAP 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   1035
      End
      Begin VB.TextBox txtFuncAP 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   420
         Width           =   4635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   15
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3270
         TabIndex        =   12
         Top             =   120
         Width           =   120
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   13
      Top             =   3780
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
            Text            =   "Online.Info"
            TextSave        =   "Online.Info"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "12:32"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox mskData 
      Height          =   315
      Left            =   2520
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin ChamaleonBtn.chameleonButton cmdReativar 
      Height          =   675
      Left            =   60
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1191
      BTYPE           =   3
      TX              =   "&Abrir Caixa"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Caixa_Fechamento.frx":23D2
      PICN            =   "Caixa_Fechamento.frx":23EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdFecharCaixa 
      Height          =   675
      Left            =   3720
      TabIndex        =   6
      Top             =   3000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1191
      BTYPE           =   3
      TX              =   "&Fechar Caixa"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Caixa_Fechamento.frx":2855
      PICN            =   "Caixa_Fechamento.frx":2871
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "Caixa_Fechamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim COD_CAIXADIA As Long
Dim COD_CAIXA As Long
Private Sub AutoNumeracao_Caixa()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   COD_CAIXA = 1
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod FROM caixa_saldo;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then COD_CAIXA = r("cod") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cmdReativar_Click()
Dim var_CodCaixa As Long
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodFuncAP.Text = "" Then MsgBox "Coloque o Cód de Funcinario!", vbInformation, "Aviso do Sistema": txtCodFuncAP.SetFocus: Exit Sub

If cmdReativar.Caption = "Abrir Caixa" Then
   'pegar o código do ultimo pedido
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_caixa FROM caixa_dia"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then var_CodCaixa = r("ultimo_caixa") + 1
   
   dbData.Execute "INSERT INTO caixa_dia (codigo, data_abertura, hora_abertura, cod_funcionario, status, entrada, saida, saldo, maquina) VALUES (" & var_CodCaixa & ", CONVERT(DATETIME, '" & Format(StatusBar1.Panels(4).Text, ocDATA) & "', 103), '" & Format(Now, ocHRMN) & "', " & txtCodFuncAP.Text & ", 0, " & Replace(CCur(lblEntrada), ",", ".") & ", " & Replace(CCur(lblSaida), ",", ".") & ", " & Replace(CCur(lblTotal), ",", ".") & ", '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "');"
Else
   dbData.Execute "UPDATE caixa_dia SET status = 0 WHERE (data_abertura = CONVERT(DATETIME, '" & Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, ocDATA) & "', 103)) AND (maquina = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "');"
   ''execSQL "DELETE FROM CAIXA_DIA WHERE DATA_ABERTURA = #" & Format(StatusBar1.Panels(3).Text, "dd/mm/yyyy") & "# and CAIXA = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "' "
   'dbData.Execute "DELETE FROM caixa_saldo WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103));"
End If

Unload Me
End Sub

Private Sub cmdFecharCaixa_Click()
   'On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim ANTERIOR As Currency
   
   If txtFuncAP.Text = "" Then
      ShowMsg "Faltou o código do funcionário!", vbExclamation
      txtCodFuncAP.SetFocus
      Exit Sub
   End If
   
   If lblEntrada.Caption = "" Or lblSaida.Caption = "" Or lblTotal.Caption = "" Then
      Exit Sub
   End If
   
   'AUTONUMERACAO CAIXA_DIA
   'Autonumeracao_CaixaDia
   
   'SALVAR NA TABELA CAIXA_DIA
   'execSQL "INSERT INTO CAIXA_DIA (CODIGO, Data, ENTRADA, SAIDA, SALDO, Status) VALUES(" & COD_CAIXADIA & ", #" & Format(mskData.Text, "mm/dd/yyyy") & "#, '" & lblEntrada.Caption & "', '" & lblSaida.Caption & "', '" & lblTotal.Caption & "', true)"
   sSQL = "UPDATE caixa_dia SET " & _
      "entrada = " & Replace(CCur(lblEntrada), ",", ".") & ", " & _
      "saida = " & Replace(CCur(lblSaida), ",", ".") & ", " & _
      "saldo = " & Replace(CCur(lblTotal), ",", ".") & ", " & _
      "status = 1, " & _
      "data_fechamento = CONVERT(DATETIME, '" & Format(StatusBar1.Panels(3).Text, ocDATA) & "', 103), " & _
      "hora_fechamento = '" & Format(Now, ocHRMN) & "', " & _
      "cod_funcionario = " & txtCodFuncAP & _
      " WHERE (codigo = " & txtCodCaixa.Text & ");"
    Debug.Print sSQL
   dbData.Execute sSQL
   
   'SALVAR NA TABELA CAIXA====================================
   
   'função para saber o saldo anterior
   sSQL = "SELECT TOP 1 ISNULL(saldo_atual, 0) AS saldo_caixa FROM caixa_saldo ORDER BY data DESC;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then ANTERIOR = r("saldo_caixa")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'autonumeração
   AutoNumeracao_Caixa
   
   'Salvar Caixa
   dbData.Execute "INSERT INTO caixa_saldo (codigo, data, saldo_anterior, entrada, retirada, saldo_atual) VALUES(" & _
      COD_CAIXA & ", CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103), " & Replace(ANTERIOR, ",", ".") & ", " & _
      Replace(CCur(lblTotal.Caption), ",", ".") & ", 0, " & Replace((ANTERIOR + CCur(lblTotal.Caption)), ",", ".") & ");"
   
   'Form_Load
   Unload Me
   Exit Sub
   
'TrataErro:
   'If Err.Number = 3022 Then
   '   ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
   '   Exit Sub
   'End If
End Sub

Private Sub AutoNumeracao_CaixaDia()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   COD_CAIXADIA = 0
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_caixa FROM caixa_dia;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then COD_CAIXADIA = r("cod_caixa") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Form_Load()
   Dim var_Entrada As Currency, var_Saida As Currency, var_Total As Currency
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim var_Maquina As String     'colocar o nome da maquina na barra de status
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
   Set oIni = Nothing
   
   StatusBar1.Panels(2).Text = var_Maquina
   
   'txtCodCaixa.Text = Caixa_Controle_semOS.txtCodCaixa.Text 'ver depois
    StatusBar1.Panels(4).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yy")
    mskData.Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yy")

   'mskData.Text = Format(Date, "dd/mm/yy")
   lblEntrada.Caption = Caixa_Controle_semOS.txtEntrada.Text
   lblSaida.Caption = Caixa_Controle_semOS.txtSaida.Text
   
   var_Entrada = lblEntrada.Caption
   var_Saida = lblSaida.Caption
   var_Total = var_Entrada - var_Saida
   lblTotal.Caption = Format(var_Total, ocMONEY)
   
   'MOSTRAR SE O CAIXA ESTÁ FECHADO
   sSQL = "SELECT codigo, data_abertura, maquina, status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, ocDATA) & "', 103)) AND (maquina = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "');"
   Set r = dbData.OpenRecordset(sSQL)

    

   If r.BOF Then
      cmdReativar.Caption = "Abrir Caixa"
      cmdReativar.Enabled = True       'tem que apagar essa linha depois
      cmdFecharCaixa.Enabled = False
   Else
      If CInt(ValidateNull(r("status"))) = 0 Then
         txtCodCaixa.Text = CInt(ValidateNull(r("codigo")))
         cmdReativar.Enabled = False
         cmdFecharCaixa.Enabled = True
      Else
         txtCodCaixa.Text = CInt(ValidateNull(r("codigo")))
         cmdReativar.Enabled = True
         cmdReativar.Caption = "Reativar Caixa"
         cmdFecharCaixa.Enabled = False
      End If
   End If
   
   'MOSTRAR SE O CAIXA ESTÁ FECHADO
  ' sSQL = "SELECT TOP 1 * FROM caixa_dia order by codigo desc;"
  ' Set r = dbData.OpenRecordset(sSQL)
   
  '  If Not r.EOF Then
  '      If r("status") = True Then
  '          cmdReativar.Caption = "Abrir Caixa"
  '          cmdReativar.Enabled = True       'tem que apagar essa linha depois
  '          cmdFecharCaixa.Enabled = False
  '      Else
  '          cmdReativar.Enabled = False
  '          cmdFecharCaixa.Enabled = True
  '      End If
  '  Else
  '      StatusBar1.Panels(2).Text = "FECHADO"  'caso nao exista registro na tabela
  '  End If
      
      'If Tela_Principal.StatusBar1.Panels(2).Text = "GERENTE" Then cmdReativar.Enabled = True Else cmdReativar.Enabled = False
   
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub lblSaida_Change()
   '=============TOTAL DA SAIDA===============
   Dim TotalTE As Currency
   Dim TotalTS As Currency
   Dim TotalTT As Currency
   
   TotalTE = lblEntrada.Caption
   TotalTS = lblSaida.Caption
   TotalTT = TotalTE - TotalTS
   lblTotal.Caption = Format(TotalTT, ocMONEY)
End Sub

Private Sub txtCodFuncAP_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodFuncAP.Text = "" Then Exit Sub
   txtFuncAP.Text = ""
   
   sSQL = "SELECT codigo, nome, sobrenome FROM funcionario WHERE (codigo = " & txtCodFuncAP.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtFuncAP.Text = r("nome") & " " & r("sobrenome")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub txtCodFuncAP_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub
