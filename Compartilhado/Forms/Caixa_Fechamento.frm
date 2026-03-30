VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Caixa_Fechamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SITUAÇÃO DO CAIXA"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "Caixa_Fechamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   60
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   2595
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
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   1500
      Visible         =   0   'False
      Width           =   3075
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lblSaida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         TabIndex        =   5
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         TabIndex        =   6
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label lblEntrada 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         TabIndex        =   4
         Top             =   540
         Width           =   1875
      End
   End
   Begin VB.Frame Frame7 
      Height          =   915
      Left            =   60
      TabIndex        =   8
      Top             =   540
      Width           =   7695
      Begin VB.TextBox mskData 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox txtTroco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   1
         Top             =   420
         Width           =   1515
      End
      Begin VB.TextBox txtHora 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txtCodFuncAP 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txtFuncAP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   420
         Width           =   2655
      End
      Begin VB.Label lblTroco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Troco"
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
         Left            =   5880
         TabIndex        =   25
         Top             =   180
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
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
         TabIndex        =   24
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora"
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
         Left            =   1380
         TabIndex        =   23
         Top             =   180
         Width           =   525
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
         Left            =   2280
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
         Left            =   3180
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
         TabIndex        =   13
         Top             =   120
         Width           =   120
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   14
      Top             =   1980
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7382
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
            TextSave        =   "18:12"
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
   Begin ChamaleonBtn.chameleonButton cmdReativar 
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   1500
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
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
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   1500
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
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
      MICON           =   "Caixa_Fechamento.frx":23EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "ABERTURA DO CAIXA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   7635
   End
End
Attribute VB_Name = "Caixa_Fechamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim COD_CAIXADIA As Long
Private printSQL As String
'Dim varCodSaldo As Long     'Usado no saldo
Dim sSQL As String
Dim r As ADODB.Recordset

Dim ANTERIOR As Currency    'usado no saldo para saber o saldo do ultimo caixa
Dim varCodSaldo As Integer  'usado no saldo para saber o codigo do ultimo caixa antes do update
Dim vOSAtiva As Boolean
Dim vAluguelAtiva As Boolean

Public vConfImprimeNFCeLocal As String

'arquivo .ini
Public cCfg As ConfigItem
Public oIni As Ini
Private Sub AutoNumeracao_Saldos()
Dim sSQL As String
Dim r As ADODB.Recordset

varCodSaldo = 1
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod FROM caixa_saldo;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then varCodSaldo = r("cod") + 1
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Backup()
Dim rEmpresa As ADODB.Recordset, xCaminhoBK As String
Dim NomeEmp As String, i As Integer, ComandoSQL As String, e As String, nomeArquivoBK As String

'parte de encontrar o caminho do sistema
sSQL = "SELECT DiretorioXML, razao, CNPJ FROM Empresa"
Set rEmpresa = dbData.OpenRecordset(sSQL)

If Not rEmpresa.EOF Then
    dirXML = IIf(Right(rEmpresa!DiretorioXML, 1) = "\", rEmpresa!DiretorioXML, rEmpresa!DiretorioXML & "\")
End If

xCaminhoBK = dirXML & "backup"

'cria a pasta caso não exista
If Not Existe(xCaminhoBK) Then MkDir xCaminhoBK

If Not Existe(xCaminhoBK) Then Exit Sub

'nomeArquivoBK = Format(Date, "yyyy-mm-dd") & "__" & rEmpresa!Razao & ".bak"
nomeArquivoBK = Retira(rEmpresa!CNPJ, ".-/ ", UM_A_UM) & ".bak"
DoEvents

If Dir$(xCaminhoBK & "\" & nomeArquivoBK) <> "" Then
   Kill xCaminhoBK & "\" & nomeArquivoBK
   Do While Dir$(xCaminhoBK & "\" & nomeArquivoBK) <> ""
      Sleep (200)
   Loop
End If

ComandoSQL = "EXEC BackupBD '" & xCaminhoBK & "'"
e$ = SQLExecuta(ComandoSQL)
If e$ <> "" Then
   MsgBox e$, vbCritical + vbOKOnly, "ERRO BACKUP"
   Exit Sub
End If

Do While Dir$(xCaminhoBK & "\" & nomeArquivoBK) = ""
   Sleep (200)
Loop


If Dir$(xCaminhoBK & "\" & nomeArquivoBK) <> "" Then
   iRetorno = CompactarBackup2(nomeArquivoBK, xCaminhoBK)
   On Error Resume Next
   Sleep (1000)
   If iRetorno Then
      Kill xCaminhoBK & "\" & nomeArquivoBK
   End If
End If
End Sub
Private Function CompactarBackup2(nomeArquivoBK As String, diretorioDestino As String) As Boolean
Dim FileZipName As String, PathToCompress As String, DestPath As String, FullPathZip As String
Dim NomeEmp As String, emailDestino As String, i As Integer, ComandoSQL As String
Dim rsEntradas As New ADODB.Recordset, rsNFe As New ADODB.Recordset, rsNFCe As New ADODB.Recordset
Dim xDiretorioDestino As String, xArquivoDestino As String

On Error GoTo deuErro

'INICIAR COMPACTAÇÃO
If IniciaComponenteCompactacao Then

   i = 0
   
   'Caminho para comprimir arquivo
   DestPath = diretorioDestino
   
   'nome do arquivo
   'NomeEmp = vRazao
   'NomeEmp = RemoveAcento(NomeEmp)
   'NomeEmp = Substitui(NomeEmp, ".,/", "", UM_A_UM)
   'NomeEmp = Substitui(NomeEmp, " ", "_", UM_A_UM)
   FileZipName = Left(nomeArquivoBK, Len(nomeArquivoBK) - 4) & ".rar"
   
   If Dir$(diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & FileZipName) <> "" Then
      Kill diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & FileZipName
      DoEvents
      Do While Dir$(diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & FileZipName) <> ""
         Sleep (200)
      Loop
   End If
   
   'local de destino + ficheiro.rar
   'diretorioDestino = vCaminhoXML & "\backup"
   FullPathZip = Transforma(diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & FileZipName)
   
   'DiretorioDestino = vCaminhoXML & "\nfe\arquivos\procNFe" & "\" & vAno & vMes
   'DiretorioOrigem = nomeArquivoBK  'vCaminhoXML & "\nfe\arquivos\procNFe" & "\" & vAno & vMesNum
     
   If Not Existe(diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & nomeArquivoBK) Then MsgBox "Não existe o arquivo de BACKUP informado!", vbInformation, "Aviso do Sistema": Exit Function
   
   PathToCompress = Transforma(diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & nomeArquivoBK)
   
   'Chama o compressor que se encontra instalado para o efeito.
   If xWinRar <> "" Then
       Shell xWinRar & " a -ep1 " & FullPathZip & " " & PathToCompress, vbNormalFocus  ', vbHide
   Else
       Shell xWinZip & " -a -ep1 " & FullPathZip & " " & PathToCompress, vbNormalFocus ', vbHide
   End If
   
   DoEvents
   
   'Load frmECFMsg
   'frmECFMsg.SetaMensagem "Aguarde! Compactando BACKUP..."
  
   'entra em loop até a criação do arquivo
   Do While Dir$(diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & FileZipName) = ""
       Sleep (200)
   Loop
   
   'Unload frmECFMsg
   
End If

CompactarBackup2 = True
Exit Function

deuErro:
  CompactarBackup2 = False
  ''If FormExists("frmECFMsg") Then Unload frmECFMsg
'lblAguarde.Visible = False
End Function

Private Sub GerarSaldo()
Dim sSQLSaldo As String
Dim rSaldo As ADODB.Recordset

sSQLSaldo = "SELECT * FROM caixa_saldo WHERE (codcaixa = " & txtCodCaixa.Text & ") and (caixa = '" & StatusBar1.Panels(2).Text & "');"
Set rSaldo = dbData.OpenRecordset(sSQLSaldo)


'criar ou modificar saldo
If Not rSaldo.BOF Then
    varCodSaldo = rSaldo("codigo")
    SaberSaldoAnteriorUpdate
    
    dbData.Execute "UPDATE caixa_saldo SET " & _
    "codigo = " & varCodSaldo & ", " & _
    "data = CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103), " & _
    "saldo_anterior = " & Replace(ANTERIOR, ",", ".") & ", " & _
    "entrada = " & Replace(CCur(lblTotal.Caption), ",", ".") & ", " & _
    "retirada = 0, " & _
    "saldo_atual = " & Replace((ANTERIOR + CCur(lblTotal.Caption)), ",", ".") & " " & _
    "WHERE codigo = " & varCodSaldo & ";"

Else
    SaberSaldoAnterior
    
    AutoNumeracao_Saldos
    
    dbData.Execute "INSERT INTO caixa_saldo (codigo, data, saldo_anterior, entrada, retirada, saldo_atual, codcaixa, caixa) VALUES(" & _
    varCodSaldo & ", CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103), " & Replace(ANTERIOR, ",", ".") & ", " & _
    Replace(CCur(lblTotal.Caption), ",", ".") & ", 0, " & Replace((ANTERIOR + CCur(lblTotal.Caption)), ",", ".") & ", " & txtCodCaixa.Text & ", '" & StatusBar1.Panels(2).Text & "');"
End If
End Sub

Private Sub ImprimirCaixaDetalhado()
Dim sSQL As String
Dim r As ADODB.Recordset

Dim SETOR_CAIXA As String
Dim var_Setor As String
Dim varTipoCartao2 As String

'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = var_ImpNormal Then
      Set Printer = Prt
      Exit For
   End If
Next

'================ PREENCHER O GRID
        Dim Maquina_Parcela As String
        Maquina_Parcela = "AND (parcelas.caixa = '" & StatusBar1.Panels(2).Text & "') "
        
        Dim Maquina_Haver As String
        Maquina_Haver = "AND (parcelas_haver.caixa = '" & StatusBar1.Panels(2).Text & "') "
        
        Dim Maquina_Suprimento As String
        Maquina_Suprimento = "AND (caixa_entrada.caixa = '" & StatusBar1.Panels(2).Text & "') "
        
        Dim Maquina_Sangria As String
        Maquina_Sangria = "AND (caixa_saida.caixa = '" & StatusBar1.Panels(2).Text & "') "
        
        SETOR_CAIXA = "AND (pedidos.tipo_pedido = 'VENDA') "
        
        var_Setor = "AND (setor <> 'BOSTA') "
        
        sSQL = "SELECT " & _
           "parcelas.tipo as varTipoLanc, " & _
           "parcelas.hora AS varHora, " & _
           "parcelas.codigo AS varCodigo, " & _
           "pedidos.cod_pedido AS varCodPedido, " & _
           "cliente.nome AS varCliente, " & _
           "parcelas.forma_pgto AS varFormaPgto, " & _
           "parcelas.valor_final AS varValorLanc, " & _
           "0 AS varValorSaida, " & _
           "(parcelas.valor_final  - 0) AS campo04, " & _
           "(CASE WHEN parcelas.tipo_cartao = 'D' THEN 'DÉBITO' WHEN parcelas.tipo_cartao = 'C'  THEN 'CRÉDITO' Else '' End) AS varTipoCartao, " & _
           "parcelas.pagamento AS data, " & _
           "pedidos.cod_pedido AS pedido, " & _
           "pedidos.cod_cliente AS cliente, " & _
           "'' AS setor, " & _
           "parcelas.caixa " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _
           "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") " & Maquina_Parcela & _
           "UNION ALL "
        
        'parcelas_haveres
        sSQL = sSQL & "SELECT " & _
           "'HAVER' as varTipoLanc, " & _
           "parcelas_haver.hora AS varHora, " & _
           "0 AS varCodigo, " & _
           "parcelas.cod_pedido AS varCodPedido, " & _
           "(SELECT nome FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente WHERE (pedidos.cod_pedido = parcelas.cod_pedido)) AS varCliente, " & _
           "parcelas_haver.forma_pgto AS varFormaPgto, " & _
           "parcelas_haver.valor_haver AS varValorLanc, " & _
           "0 AS varValorSaida, " & _
           "parcelas_haver.valor_haver AS campo04, " & _
           "(CASE WHEN parcelas_haver.tipo_cartao = 'D' THEN 'DÉBITO' WHEN parcelas_haver.tipo_cartao = 'C'  THEN 'CRÉDITO' Else '' End) AS varTipoCartao, " & _
           "parcelas_haver.haver AS data, " & _
           "parcelas_haver.codigo AS campotc, " & _
           "0 AS cliente, " & _
           "'' AS  setor, " & _
           "'' AS maquina " & _
           "FROM parcelas_haver INNER JOIN parcelas ON parcelas_haver.cod_parcela = parcelas.codigo " & _
           "WHERE (parcelas_haver.codcaixa = " & txtCodCaixa.Text & ") " & Maquina_Haver & _
           "UNION ALL "
        
        'suprimentos
        sSQL = sSQL & "SELECT " & _
           "'SUPRIMENTO' as varTipoLanc, " & _
           "hora AS varHora, " & _
           "0 AS varCodigo, " & _
           "0 AS varCodPedido, " & _
           "descricao AS varCliente, " & _
           "caixa_entrada.forma_pgto AS varFormaPgto, " & _
           "valor AS varValorLanc, " & _
           "0 AS varValorSaida, " & _
           "valor AS campo04, " & _
           "'' AS varTipoCartao, " & _
           "data, " & _
           "0 AS pedido, " & _
           "0 AS cliente, " & _
           "setor, " & _
           " '' AS maquina " & _
           "FROM caixa_entrada WHERE (caixa_entrada.codcaixa = " & txtCodCaixa.Text & ")  " & Maquina_Suprimento & var_Setor & _
           "UNION ALL "
           
        'sangria
        sSQL = sSQL & "SELECT " & _
           "'SANGRIA' as varTipoLanc, " & _
           "caixa_saida.hora AS varHora, " & _
           "0 AS varCodigo, " & _
           "0 AS varCodPedido, " & _
           "caixa_saida.descricao AS varCliente, " & _
           "'DINHEIRO' AS varFormaPgto, " & _
           "0 AS varValorLanc, " & _
           "caixa_saida.valor AS varValorSaida, " & _
           "(0 - caixa_saida.valor) AS campo04, " & _
           "'' AS varTipoCartao, " & _
           "data, " & _
           "0, " & _
           "0, " & _
           "setor, " & _
           "'' AS maquina " & _
           "FROM caixa_saida WHERE (caixa_saida.codcaixa = " & txtCodCaixa.Text & ")  " & Maquina_Sangria & var_Setor & _
           " ORDER BY 2"
        
        Set r = dbData.OpenRecordset(sSQL)
        
        printSQL = sSQL
        
        '=========== DEFINIR A IMPRESSORA
        'Dim var_Impressora As String
        'Dim oIni As Ini
        
        'Set oIni = New Ini
        'oIni.Arquivo = appPathApp & "config.ini"
        'var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
        'Set oIni = Nothing
        
        Me.Hide


Set r = dbData.OpenRecordset(printSQL)

        If lblEntrada.Caption = "0,00" Then   'fiz esse if para imprimir caixa sem saldo
            If r.State <> 0 Then r.Close
            Set r = Nothing
        End If


''==================== PEGAR OS DADOS DO FECHAMENTO
Dim sSQLusuario As String
Dim r_usuario As ADODB.Recordset

sSQLusuario = "SELECT DATA_ABERTURA, HORA_ABERTURA, COD_FUNC_ABERTURA, DATA_FECHAMENTO, HORA_FECHAMENTO, COD_FUNC_FECHAMENTO, (CASE WHEN status = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS VarStatus, " & _
        "(SELECT Usuario.Login FROM Usuario INNER JOIN caixa_dia ON Usuario.Codigo = caixa_dia.COD_FUNC_ABERTURA wHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ")) AS Nome_Abertura, " & _
        "(SELECT Usuario_2.Login FROM Usuario AS Usuario_2 INNER JOIN caixa_dia AS caixa_dia_2 ON Usuario_2.Codigo = caixa_dia_2.COD_FUNC_FECHAMENTO WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ")) AS Nome_Fechamento " & _
       "FROM caixa_dia AS caixa_dia_1 " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ");"
Set r_usuario = dbData.OpenRecordset(sSQLusuario)

'SETAR O RELATORIO
Set REL_Caixa_Fech_Imp_Todos.ReportMain1.Recordset = r

'==================== CABEÇALHO
If Not r_usuario.EOF Then
    REL_Caixa_Fech_Imp_Todos.txtDHead.Caption = "FECHAMENTO DE CAIXA - ABERTURA: " & Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
End If

'===========================CALCULO DOS TOTAIS

'VENDAS DINHEIRO============
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaVendas " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'VENDA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalDinheiroVenda As Currency
If Not r.EOF Then
    varTotalDinheiroVenda = Format(ValidateNull(r("varSomaVendas")), "#,##0.00")
Else
    varTotalDinheiroVenda = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp_Todos.rfDinheiro.Caption = Format(varTotalDinheiroVenda, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'VENDA')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfDinheiroQuant.Caption = Format(r.RecordCount, "000") & " "


'PARCELAS DINHEIRO============
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaParcelas " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'PARCELA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalDinheiroParcela As Currency
If Not r.EOF Then
    varTotalDinheiroParcela = Format(ValidateNull(r("varSomaParcelas")), "#,##0.00")
Else
    varTotalDinheiroParcela = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp_Todos.rfParcelas.Caption = Format(varTotalDinheiroParcela, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'PARCELA')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfParcelasQuant.Caption = Format(r.RecordCount, "000") & " "

'PARCELAS HAVER============
sSQL = "SELECT SUM(VALOR_HAVER) AS varSomaHaveres " & _
       "FROM parcelas_haver " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DINHEIRO')"
Set r = dbData.OpenRecordset(sSQL)

Dim varValorHaveres As Currency
If Not r.EOF Then
    varValorHaveres = Format(ValidateNull(r("varSomaHaveres")), "#,##0.00")
Else
    varValorHaveres = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp_Todos.rfHaveres.Caption = Format(varValorHaveres, "#,##0.00") & " "

sSQL = "SELECT CODIGO " & _
       "FROM parcelas_haver " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DINHEIRO')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfHaveresQuant.Caption = Format(r.RecordCount, "000") & " "




'ALUGUEL DINHEIRO============
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaAluguel " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'ALUGUEL')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalDinheiroAluguel As Currency
If Not r.EOF Then
    varTotalDinheiroAluguel = Format(ValidateNull(r("varSomaAluguel")), "#,##0.00")
Else
    varTotalDinheiroAluguel = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp_Todos.rfAluguel.Caption = Format(varTotalDinheiroAluguel, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'ALUGUEL')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfAluguelQuant.Caption = Format(r.RecordCount, "000") & " "


'OS DINHEIRO============
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaOS " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'OS')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalDinheiroOS As Currency
If Not r.EOF Then
    varTotalDinheiroOS = Format(ValidateNull(r("varSomaOS")), "#,##0.00")
Else
    varTotalDinheiroOS = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp_Todos.rfOS.Caption = Format(varTotalDinheiroOS, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DINHEIRO') AND (TIPO = 'OS')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfOSQuant.Caption = Format(r.RecordCount, "000") & " "







'CARTÃO============
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaCartao1 " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'CARTAO')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalCartao As Currency
varTotalCartao = Format(ValidateNull(r("varSomaCartao1")))

sSQL = "SELECT SUM(VALOR_HAVER) AS varSomaCartao2 " & _
       "FROM parcelas_HAVER " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'CARTAO') "
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalCartao2 As Currency
varTotalCartao2 = Format(ValidateNull(r("varSomaCartao2")))

varTotalCartao = varTotalCartao + varTotalCartao2

'If Not r.EOF Then
'    varTotalCartao = ValidateNull(r("varTotalCartao"))
'Else
'    varTotalCartao = Format(0, "#,##0.00")
'End If

REL_Caixa_Fech_Imp_Todos.rfCartao.Caption = Format(varTotalCartao, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'CARTAO')"
Set r = dbData.OpenRecordset(sSQL)

Dim ContaCartao1 As Integer
Dim ContaCartao2 As Integer
ContaCartao1 = r.RecordCount

sSQL = "SELECT codigo " & _
       "FROM parcelas_haver " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'CARTAO')"
Set r = dbData.OpenRecordset(sSQL)

ContaCartao2 = ContaCartao1 + r.RecordCount

REL_Caixa_Fech_Imp_Todos.rfCartaoQuant.Caption = Format(ContaCartao2, "000") & " "

'CHEQUE============

sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaCheque " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'CHEQUE')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalCheque As Currency

If Not r.EOF Then
    varTotalCheque = Format(ValidateNull(r("varSomaCheque")), "#,##0.00")
Else
    varTotalCheque = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp_Todos.rfCheque.Caption = Format(varTotalCheque, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'CHEQUE')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfChequeQuant.Caption = Format(r.RecordCount, "000") & " "

'DEPOSITO/TRANSFERENCIA/BOLETO/FINANCEIRA============
'boleto
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaBoleto " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'BOLETO')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalBoleto As Currency
If Not r.EOF Then
    varTotalBoleto = ValidateNull(r("varSomaBoleto"))
Else
    varTotalBoleto = 0
End If

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'BOLETO')"
Set r = dbData.OpenRecordset(sSQL)

Dim contaBoleto As Integer
contaBoleto = Format(r.RecordCount, "000")

'transferencia
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaTransferencia " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'TRANSFERENCIA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalTransferencia As Currency
If Not r.EOF Then
    varTotalTransferencia = ValidateNull(r("varSomaTransferencia"))
Else
    varTotalTransferencia = 0
End If

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'TRANSFERENCIA')"
Set r = dbData.OpenRecordset(sSQL)

Dim contaTransferencia As Integer
contaTransferencia = Format(r.RecordCount, "000")

'deposito
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaDeposito " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DEPOSITO')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalDeposito As Currency
If Not r.EOF Then
    varTotalDeposito = ValidateNull(r("varSomaDeposito"))
Else
    varTotalDeposito = 0
End If

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'DEPOSITO')"
Set r = dbData.OpenRecordset(sSQL)

Dim contaDeposito As Integer
contaDeposito = Format(r.RecordCount, "000")

'financeira
sSQL = "SELECT SUM(VALOR_FINAL) AS varSomaFinanceira " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'FINANCEIRA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalFinanceira As Currency
If Not r.EOF Then
    varTotalFinanceira = ValidateNull(r("varSomaFinanceira"))
Else
    varTotalFinanceira = 0
End If

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (FORMA_PGTO = 'FINANCEIRA')"
Set r = dbData.OpenRecordset(sSQL)

'soma
Dim varTotalBTD As Currency
varTotalBTD = varTotalBoleto + varTotalTransferencia + varTotalDeposito + varTotalFinanceira

REL_Caixa_Fech_Imp_Todos.rfOutros.Caption = Format(varTotalBTD, "#,##0.00") & " "

Dim contaFinanceira As Integer
contaFinanceira = Format(r.RecordCount, "000")

Dim ContaOutros As Integer
ContaOutros = contaFinanceira + contaDeposito + contaTransferencia + contaBoleto

REL_Caixa_Fech_Imp_Todos.rfOutrosQuant.Caption = Format(ContaOutros, "000") & " "

'SANGRIA============
sSQL = "SELECT SUM(VALOR) AS varSomaSangria " & _
       "FROM caixa_saida " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") and (FONTE = 'CAIXA ATUAL') "
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalSangria As Currency

If Not r.EOF Then
    varTotalSangria = Format(ValidateNull(r("varSomaSangria")), "#,##0.00")
Else
    varTotalSangria = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp_Todos.rfSaida.Caption = Format(varTotalSangria, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM caixa_saida " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") and (FONTE = 'CAIXA ATUAL') "
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfSaidaQuant.Caption = Format(r.RecordCount, "000") & " "

'SUPRIMENTO============
sSQL = "SELECT SUM(VALOR) AS varSomaSuprimento " & _
       "FROM caixa_entrada " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") "
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalSuprimento As Currency

If Not r.EOF Then
    varTotalSuprimento = Format(ValidateNull(r("varSomaSuprimento")), "#,##0.00")
Else
    varTotalSuprimento = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp_Todos.rfSuprimentos.Caption = Format(varTotalSuprimento, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM caixa_entrada " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") "
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfSuprimentosQuant.Caption = Format(r.RecordCount, "000") & " "

'TROCO============
sSQL = "SELECT SUM(VALOR) AS varSomaTROCO " & _
       "FROM caixa_troco " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") "
Set r = dbData.OpenRecordset(sSQL)

'If Not r.EOF Then
'    REL_Caixa_Fech_Imp_Todos.rfTroco.Caption = Format(ValidateNull(r("varSomaTROCO")), "#,##0.00") & " "
'Else
'    REL_Caixa_Fech_Imp_Todos.rfTroco.Caption = Format(0, "#,##0.00") & " "
'End If

'VENDA A PRAZO ================
sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaPrazo " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & txtCodCaixa.Text & ") AND pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (parcelas.STATUS = 0)"
'Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalPrazo As Currency

If Not r.EOF Then
    varTotalPrazo = ValidateNull(r("varSomaPrazo"))
Else
    varTotalPrazo = Format(0, "#,##0.00")
End If

REL_Caixa_Fech_Imp_Todos.rfPrazo.Caption = Format(varTotalPrazo, "#,##0.00") & " "

sSQL = "SELECT parcelas.cod_pedido " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & txtCodCaixa.Text & ") AND pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (parcelas.STATUS = 0)"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfPrazoQuant.Caption = Format(r.RecordCount, "000") & " "

'CALCULAR TOTAIS================
Dim varTotaisEntrada As Currency
Dim varTotaisSaida As Currency
varTotaisEntrada = varTotalDinheiroVenda + varTotalDinheiroParcela + varValorHaveres + varTotalCheque + varTotalSuprimento + varTotalDinheiroAluguel + varTotalDinheiroOS
varTotaisSaida = varTotaisEntrada - varTotalSangria

REL_Caixa_Fech_Imp_Todos.rfSaldoFisico.Caption = Format(varTotaisSaida, "#,##0.00") & " "

Dim varTotaisGeral As Currency
varTotaisGeral = varTotaisSaida + varTotalBTD + varTotalCartao

REL_Caixa_Fech_Imp_Todos.rfSaldoGeral.Caption = Format(varTotaisGeral, "#,##0.00") & " "

Dim varTotaisFAT As Currency
varTotaisFAT = varTotaisGeral + varTotalPrazo
REL_Caixa_Fech_Imp_Todos.rfFaturamento.Caption = Format(varTotaisFAT, "#,##0.00") & " "

'===========================RODAPÉ
If Not r_usuario.EOF Then
    REL_Caixa_Fech_Imp_Todos.rfCodUsuarioA.Caption = Format(r_usuario("COD_FUNC_ABERTURA"), "00")
    REL_Caixa_Fech_Imp_Todos.rfNomeUsuarioA.Caption = ValidateNull(r_usuario("Nome_Abertura"))
    REL_Caixa_Fech_Imp_Todos.rfDataA.Caption = Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
    REL_Caixa_Fech_Imp_Todos.rfHoraA.Caption = Format(ValidateNull(r_usuario("HORA_ABERTURA")), "hh:mm")
    
    REL_Caixa_Fech_Imp_Todos.rfNomeUsuarioF.Caption = ValidateNull(r_usuario("Nome_Fechamento"))
    If IsNull(r_usuario("DATA_FECHAMENTO")) Then
        REL_Caixa_Fech_Imp_Todos.rfDataF.Caption = ""
        REL_Caixa_Fech_Imp_Todos.rfCodUsuarioF.Caption = ""
        REL_Caixa_Fech_Imp_Todos.rfHoraF.Caption = ""
    Else
        REL_Caixa_Fech_Imp_Todos.rfCodUsuarioF.Caption = Format(ValidateNull(r_usuario("COD_FUNC_FECHAMENTO")), "00")
        REL_Caixa_Fech_Imp_Todos.rfDataF.Caption = Format(ValidateNull(r_usuario("DATA_FECHAMENTO")), "dd/mm/yyyy")
        REL_Caixa_Fech_Imp_Todos.rfHoraF.Caption = Format(ValidateNull(r_usuario("HORA_FECHAMENTO")), "hh:mm")
    End If

    REL_Caixa_Fech_Imp_Todos.rfSituacao.Caption = ValidateNull(r_usuario("VARSTATUS"))
End If

REL_Caixa_Fech_Imp_Todos.rfCaixa.Caption = StatusBar1.Panels(2).Text
REL_Caixa_Fech_Imp_Todos.rfCodCaixa.Caption = Format(txtCodCaixa.Text, "0000")


'=========================CALCULO DO FATURAMENTO

'VENDAS
sSQL = "SELECT ISNULL(SUM(VALOR_FINAL),0) AS varSomaVendasFAT " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (TIPO = 'VENDA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalVendaFAT As Currency
varTotalVendaFAT = ValidateNull(r("varSomaVendasFAT"))
REL_Caixa_Fech_Imp_Todos.rfT1.Caption = Format(varTotalVendaFAT, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (TIPO = 'VENDA')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfF1.Caption = Format(r.RecordCount, "000") & " "


'ALUGUEL
sSQL = "SELECT ISNULL(SUM(VALOR_FINAL),0) AS varSomaALUGUELFAT " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (TIPO = 'ALUGUEL')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalALUGUELFAT As Currency
varTotalALUGUELFAT = ValidateNull(r("varSomaALUGUELFAT"))
REL_Caixa_Fech_Imp_Todos.rfT7.Caption = Format(varTotalALUGUELFAT, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (TIPO = 'ALUGUEL')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfF7.Caption = Format(r.RecordCount, "000") & " "

'OS
sSQL = "SELECT ISNULL(SUM(VALOR_FINAL),0) AS varSomaOSFAT " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (TIPO = 'OS')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalOSFAT As Currency
varTotalOSFAT = ValidateNull(r("varSomaOSFAT"))
REL_Caixa_Fech_Imp_Todos.rfT8.Caption = Format(varTotalOSFAT, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (TIPO = 'OS')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfF8.Caption = Format(r.RecordCount, "000") & " "




'PARCELAS
sSQL = "SELECT ISNULL(SUM(VALOR_FINAL),0) AS varSomaParcelasFAT " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (TIPO = 'PARCELA')"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalParcelasFAT As Currency
varTotalParcelasFAT = ValidateNull(r("varSomaParcelasFAT"))
REL_Caixa_Fech_Imp_Todos.rfT2.Caption = Format(varTotalParcelasFAT, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") AND (TIPO = 'PARCELA')"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfF2.Caption = Format(r.RecordCount, "000") & " "

'HAVER
sSQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) AS varSomaHaveresFAT " & _
       "FROM parcelas_haver " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ")"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalHaveresFAT As Currency
varTotalHaveresFAT = ValidateNull(r("varSomaHaveresFAT"))
REL_Caixa_Fech_Imp_Todos.rfT3.Caption = Format(varTotalHaveresFAT, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM parcelas_haver " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ")"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfF3.Caption = Format(r.RecordCount, "000") & " "

'SUPRIMENTO
sSQL = "SELECT ISNULL(SUM(VALOR),0) AS varSomaSuprimentoFAT " & _
       "FROM caixa_entrada " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") "
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalSuprimentoFAT As Currency
varTotalSuprimentoFAT = ValidateNull(r("varSomaSuprimentoFAT"))
REL_Caixa_Fech_Imp_Todos.rfT4.Caption = Format(varTotalSuprimentoFAT, "#,##0.00") & " "

sSQL = "SELECT codigo " & _
       "FROM caixa_entrada " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") "
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfF4.Caption = Format(r.RecordCount, "000") & " "

'PRAZO
sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaPrazoFAT " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & txtCodCaixa.Text & ") AND pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (parcelas.STATUS = 0)"
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalPrazoFAT As Currency
varTotalPrazoFAT = ValidateNull(r("varSomaPrazoFAT"))
REL_Caixa_Fech_Imp_Todos.rfT5.Caption = Format(varTotalPrazoFAT, "#,##0.00") & " "

sSQL = "SELECT parcelas.codigo " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & txtCodCaixa.Text & ") AND pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (parcelas.STATUS = 0)"
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfF5.Caption = Format(r.RecordCount, "000") & " "

'SANGRIA
sSQL = "SELECT ISNULL(SUM(VALOR),0) AS varSomaSangriaFAT " & _
       "FROM caixa_saida " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") and (FONTE = 'CAIXA ATUAL') "
Set r = dbData.OpenRecordset(sSQL)

Dim varTotalSangriaFAT As Currency
varTotalSangriaFAT = ValidateNull(r("varSomaSangriaFAT"))
REL_Caixa_Fech_Imp_Todos.rfT6.Caption = Format(varTotalSangriaFAT, "#,##0.00") & " "

sSQL = "SELECT CODIGO " & _
       "FROM caixa_saida " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ") and (FONTE = 'CAIXA ATUAL') "
Set r = dbData.OpenRecordset(sSQL)

REL_Caixa_Fech_Imp_Todos.rfF6.Caption = Format(r.RecordCount, "000") & " "

'CALCULAR TOTAIS================
Dim varTotaisEntradaFAT As Currency
Dim varTotaisSaidaFAT As Currency

varTotaisEntradaFAT = varTotalVendaFAT + varTotalParcelasFAT + varTotalHaveresFAT + varTotalSuprimentoFAT + varTotalPrazoFAT
varTotaisSaidaFAT = varTotaisEntradaFAT - varTotalSangriaFAT

REL_Caixa_Fech_Imp_Todos.rfFTotal.Caption = Format(varTotaisSaidaFAT, "#,##0.00") & " "

REL_Caixa_Fech_Imp_Todos.ReportMain1.NomeImpressora = var_ImpNormal
REL_Caixa_Fech_Imp_Todos.ReportMain1.Ativar
Unload REL_Caixa_Fech_Imp_Todos

'Me.Show 1
End Sub

Private Sub ImprimirCaixaResumido()
Dim SETOR_CAIXA As String
Dim var_Setor As String
Dim varTipoCartao2 As String
Dim SQL As String

If Not IsDate(mskData) Then Exit Sub

'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = var_ImpNormal Then
      Set Printer = Prt
      Exit For
   End If
Next

Dim varCodCaixa As Long
varCodCaixa = txtCodCaixa.Text

If varCodCaixa = 0 Then
    SQL = "SELECT SUM(parcelas.valor_final) as vValorVendasTotal, 'VENDAS' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE 1=0"
    Set r = dbData.OpenRecordset(SQL)
Else
    Dim Maquina_Parcela As String
    If Caixa_Controle_semOS.StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Parcela = "AND (parcelas.caixa = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Parcela = "AND (parcelas.caixa <> 'CAIXA') "
    End If

    Dim Maquina_Venda As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Venda = "AND (pedidos.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Venda = "AND (pedidos.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Haver As String
    If Caixa_Controle_semOS.StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Haver = "AND (parcelas_haver.caixa = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "') "
    ElseIf Caixa_Controle_semOS.StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Haver = "AND (parcelas_haver.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Suprimento As String
    If Caixa_Controle_semOS.StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Suprimento = "AND (caixa_entrada.caixa = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "') "
    ElseIf Caixa_Controle_semOS.StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Suprimento = "AND (caixa_entrada.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Sangria As String
    If Caixa_Controle_semOS.StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Sangria = "AND (caixa_saida.caixa = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "') "
    ElseIf Caixa_Controle_semOS.StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Sangria = "AND (caixa_saida.caixa <> 'CAIXA') "
    End If
    
    SETOR_CAIXA = "AND (pedidos.tipo_pedido = 'VENDA') "
        'codcaixa = " & txtCodCaixa.Text & ") and (caixa = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "'
    
    'TROCO
    Dim vVlrTroco As Currency
    '"SELECT * FROM caixa_troco WHERE (caixa_troco.codcaixa = " & StatusBar1.Panels(3).Text & ") AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
    SQL = "SELECT * FROM caixa_troco WHERE (caixa_troco.codcaixa = " & Caixa_Controle_semOS.StatusBar1.Panels(3).Text & ") AND (caixa = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "');"
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrTroco = r("VALOR") Else vVlrTroco = 0
    
    'VENDAS
    Dim vVlrVendasTotal As Currency
    Dim vQtdeVendasTotal As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasTotal, count(codigo) as vQuantVendasTotal FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and parcelas.tipo = 'VENDA' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasTotal = r("vValorVendasTotal"): vQtdeVendasTotal = r("vQuantVendasTotal") Else vVlrVendasTotal = 0: vQtdeVendasTotal = 0

    'Detalhamento de vendas - Dinheiro
    Dim vVlrVendasDinheiro As Currency
    Dim vQtdeVendasDinheiro As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasDinheiro, count(codigo) as vQuantVendasDinheiro FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'DINHEIRO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasDinheiro = r("vValorVendasDinheiro"): vQtdeVendasDinheiro = r("vQuantVendasDinheiro") Else vVlrVendasDinheiro = 0: vQtdeVendasDinheiro = 0


    'Detalhamento de vendas - Pix
    Dim vVlrVendasPix As Currency
    Dim vQtdeVendasPix As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasPix, count(codigo) as vQuantVendasPix FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'PIX' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasPix = r("vValorVendasPix"): vQtdeVendasPix = r("vQuantVendasPix") Else vVlrVendasPix = 0: vQtdeVendasPix = 0

    'Detalhamento de vendas - Transferencia
    Dim vVlrVendasTransferencia As Currency
    Dim vQtdeVendasTransferencia As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasTransferencia, count(codigo) as vQuantVendasTransferencia FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'TRANSFERENCIA' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasTransferencia = r("vValorVendasTransferencia"): vQtdeVendasTransferencia = r("vQuantVendasTransferencia") Else vVlrVendasTransferencia = 0: vQtdeVendasTransferencia = 0

    'Detalhamento de vendas - Deposito
    Dim vVlrVendasDeposito As Currency
    Dim vQtdeVendasDeposito As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasDeposito, count(codigo) as vQuantVendasDeposito FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'DEPOSITO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasDeposito = r("vValorVendasDeposito"): vQtdeVendasDeposito = r("vQuantVendasDeposito") Else vVlrVendasDeposito = 0: vQtdeVendasDeposito = 0

    'Detalhamento de vendas - Financeira
    Dim vVlrVendasFinanceira As Currency
    Dim vQtdeVendasFinanceira As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasFinanceira, count(codigo) as vQuantVendasFinanceira FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'FINANCEIRA' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasFinanceira = r("vValorVendasFinanceira"): vQtdeVendasFinanceira = r("vQuantVendasFinanceira") Else vVlrVendasFinanceira = 0: vQtdeVendasFinanceira = 0

    'Detalhamento de vendas - Cartão
    Dim vVlrVendasCartao As Currency
    Dim vQtdeVendasCartao As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasCartao, count(codigo) as vQuantVendasCartao FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'CARTAO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasCartao = r("vValorVendasCartao"): vQtdeVendasCartao = r("vQuantVendasCartao") Else vVlrVendasCartao = 0: vQtdeVendasCartao = 0

    'Detalhamento de vendas - Cheque
    Dim vVlrVendasCheque As Currency
    Dim vQtdeVendasCheque As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasCheque, count(codigo) as vQuantVendasCheque FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'CHEQUE' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasCheque = r("vValorVendasCheque"): vQtdeVendasCheque = r("vQuantVendasCheque") Else vVlrVendasCheque = 0: vQtdeVendasCheque = 0


    'PARCELAS
    Dim vVlrParcelasTotal As Currency
    Dim vQtdeParcelasTotal As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaTotal, count(codigo) as vQuantParcelaTotal FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasTotal = r("vValorParcelaTotal"): vQtdeParcelasTotal = r("vQuantParcelaTotal") Else vVlrParcelasTotal = 0: vQtdeParcelasTotal = 0

    'Detalhamento de Parcelas - Dinheiro
    Dim vVlrParcelasDinheiro As Currency
    Dim vQtdeParcelasDinheiro As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaDinheiro, count(codigo) as vQuantParcelaDinheiro FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'DINHEIRO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasDinheiro = r("vValorParcelaDinheiro"): vQtdeParcelasDinheiro = r("vQuantParcelaDinheiro") Else vVlrParcelasDinheiro = 0: vQtdeParcelasDinheiro = 0

    'Detalhamento de Parcelas - Pix
    Dim vVlrParcelasPix As Currency
    Dim vQtdeParcelasPix As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaPix, count(codigo) as vQuantParcelaPix FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'PIX' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasPix = r("vValorParcelaPix"): vQtdeParcelasPix = r("vQuantParcelaPix") Else vVlrParcelasPix = 0: vQtdeVendasTotal = 0: vQtdeParcelasPix = 0

    'Detalhamento de Parcelas - Transferencia
    Dim vVlrParcelasTransferencia As Currency
    Dim vQtdeParcelasTransferencia As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaTransferencia, count(codigo) as vQuantParcelaTransferencia FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'TRANSFERENCIA' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasTransferencia = r("vValorParcelaTransferencia"): vQtdeParcelasTransferencia = r("vQuantParcelaTransferencia") Else vVlrParcelasTransferencia = 0: vQtdeParcelasTransferencia = 0

    'Detalhamento de Parcelas - Deposito
    Dim vVlrParcelasDeposito As Currency
    Dim vQtdeParcelasDeposito As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaDeposito, count(codigo) as vQuantParcelaDeposito FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'DEPOSITO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasDeposito = r("vValorParcelaDeposito"): vQtdeParcelasDeposito = r("vQuantParcelaDeposito") Else vVlrParcelasDeposito = 0: vQtdeParcelasDeposito = 0

    'Detalhamento de Parcelas - Financeira
    Dim vVlrParcelasFinanceira As Currency
    Dim vQtdeParcelasFinanceira As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaFinanceira, count(codigo) as vQuantParcelaFinanceira FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'FINANCEIRA' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasFinanceira = r("vValorParcelaFinanceira"): vQtdeParcelasFinanceira = r("vQuantParcelaFinanceira") Else vVlrParcelasFinanceira = 0: vQtdeParcelasFinanceira = 0

    'Detalhamento de Parcelas - Cartão
    Dim vVlrParcelasCartao As Currency
    Dim vQtdeParcelasCartao As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaCartao, count(codigo) as vQuantParcelaCartao FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'CARTAO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasCartao = r("vValorParcelaCartao"): vQtdeParcelasCartao = r("vQuantParcelaCartao") Else vVlrParcelasCartao = 0: vQtdeParcelasCartao = 0

    'Detalhamento de Parcelas - Cheque
    Dim vVlrParcelasCheque As Currency
    Dim vQtdeParcelasCheque As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaCheque, count(codigo) as vQuantParcelaCheque FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & txtCodCaixa.Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'CHEQUE' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasCheque = r("vValorParcelaCheque"): vQtdeParcelasCheque = r("vQuantParcelaCheque") Else vVlrParcelasCheque = 0: vQtdeParcelasCheque = 0


    'HAVERES
    Dim vVlrHaveresTotal As Currency
    Dim vQtdeHaveresTotal As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresTotal, count(codigo) as vQuantHaveresTotal FROM parcelas_haver " & _
           "WHERE (codcaixa = " & txtCodCaixa.Text & ") and tipo = 'PARCELA' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresTotal = r("vValorHaveresTotal"): vQtdeHaveresTotal = r("vQuantHaveresTotal") Else vVlrHaveresTotal = 0: vQtdeHaveresTotal = 0
    
    'Detalhamento de Haveres - Dinheiro
    Dim vVlrHaveresDinheiro As Currency
    Dim vQtdeHaveresDinheiro As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresDinheiro, count(codigo) as vQuantHaveresDinheiro FROM parcelas_haver " & _
           "WHERE (codcaixa = " & txtCodCaixa.Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'DINHEIRO' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresDinheiro = r("vValorHaveresDinheiro"): vQtdeHaveresDinheiro = r("vQuantHaveresDinheiro") Else vVlrHaveresDinheiro = 0: vQtdeHaveresDinheiro = 0

    'Detalhamento de Haveres - Pix
    Dim vVlrHaveresPix As Currency
    Dim vQtdeHaveresPix As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresPix, count(codigo) as vQuantHaveresPix FROM parcelas_haver " & _
           "WHERE (codcaixa = " & txtCodCaixa.Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'PIX' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresPix = r("vValorHaveresPix"): vQtdeHaveresPix = r("vQuantHaveresPix") Else vVlrHaveresPix = 0: vQtdeHaveresPix = 0

    'Detalhamento de Haveres - Transferencia
    Dim vVlrHaveresTransferencia As Currency
    Dim vQtdeHaveresTransferencia As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresTransferencia, count(codigo) as vQuantHaveresTransferencia FROM parcelas_haver " & _
           "WHERE (codcaixa = " & txtCodCaixa.Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'TRANSFERENCIA' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresTransferencia = r("vValorHaveresTransferencia"): vQtdeHaveresTransferencia = r("vQuantHaveresTransferencia") Else vVlrHaveresTransferencia = 0: vQtdeHaveresTransferencia = 0

    'Detalhamento de Haveres - Deposito
    Dim vVlrHaveresDeposito As Currency
    Dim vQtdeHaveresDeposito As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresDeposito, count(codigo) as vQuantHaveresDeposito FROM parcelas_haver " & _
           "WHERE (codcaixa = " & txtCodCaixa.Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'DEPOSITO' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresDeposito = r("vValorHaveresDeposito"): vQtdeHaveresDeposito = r("vQuantHaveresDeposito") Else vVlrHaveresDeposito = 0: vQtdeHaveresDeposito = 0

    'Detalhamento de Haveres - Financeira
    Dim vVlrHaveresFinanceira As Currency
    Dim vQtdeHaveresFinanceira As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresFinanceira, count(codigo) as vQuantHaveresFinanceira FROM parcelas_haver " & _
           "WHERE (codcaixa = " & txtCodCaixa.Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'FINANCEIRA' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresFinanceira = r("vValorHaveresFinanceira"): vQtdeHaveresFinanceira = r("vQuantHaveresFinanceira") Else vVlrHaveresFinanceira = 0: vQtdeHaveresFinanceira = 0

    'Detalhamento de Haveres - Cartão
    Dim vVlrHaveresCartao As Currency
    Dim vQtdeHaveresCartao As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresCartao, count(codigo) as vQuantHaveresCartao FROM parcelas_haver " & _
           "WHERE (codcaixa = " & txtCodCaixa.Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'CARTAO' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresCartao = r("vValorHaveresCartao"): vQtdeHaveresCartao = r("vQuantHaveresCartao") Else vVlrHaveresCartao = 0: vQtdeHaveresCartao = 0

    'Detalhamento de Haveres - Cheque
    Dim vVlrHaveresCheque As Currency
    Dim vQtdeHaveresCheque As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresCheque, count(codigo) as vQuantHaveresCheque FROM parcelas_haver " & _
           "WHERE (codcaixa = " & txtCodCaixa.Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'CHEQUE' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresCheque = r("vValorHaveresCheque"): vQtdeHaveresCheque = r("vQuantHaveresCheque") Else vVlrHaveresCheque = 0: vQtdeHaveresCheque = 0


    'SUPRIMENTO
    Dim vVlrSuprimento As Currency
    Dim vQtdeSuprimento As Integer
    SQL = "SELECT ISNULL(SUM(VALOR),0) as vValorSuprimentoTotal, count(codigo) as vQuantSuprimentoTotal FROM caixa_entrada " & _
           "WHERE (codcaixa = " & txtCodCaixa.Text & ") " & Maquina_Suprimento
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrSuprimento = r("vValorSuprimentoTotal"): vQtdeSuprimento = r("vQuantSuprimentoTotal") Else vVlrSuprimento = 0: vQtdeSuprimento = 0


    'SANGRIA
    Dim vVlrSangria As Currency
    Dim vQtdeSangria As Integer
    SQL = "SELECT ISNULL(SUM(VALOR),0) as vValorSangriaTotal, count(codigo) as vQuantSangriaTotal FROM caixa_saida " & _
           "WHERE (codcaixa = " & txtCodCaixa.Text & ") " & Maquina_Sangria
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrSangria = r("vValorSangriaTotal"): vQtdeSangria = r("vQuantSangriaTotal") Else vVlrSangria = 0: vQtdeSangria = 0

    'Debug.Print SQL

    'Set r = dbData.OpenRecordset(SQL)

    'VENDAS A PRAZO
    Dim vVlrVendasPrazoTotal As Currency
    Dim vQtdeVendasPrazoTotal As Integer
    sSQL = "SELECT ISNULL(SUM(TOTAL), 0) AS varSomaPrazoTotais, count(cod_pedido) as varQuantPrazoTotal " & _
           "FROM pedidos  " & _
           "WHERE TIPO_PAGAMENTO = 'À Prazo' and TIPO_PEDIDO= 'VENDA' and pedidos.cancelado = 0 AND (codcaixa = " & Caixa_Controle_semOS.StatusBar1.Panels(3).Text & ") " & Maquina_Venda
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vQtdeVendasPrazoTotal = r("varQuantPrazoTotal") Else vQtdeVendasPrazoTotal = 0

    sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaPrazoTotais, count(pedidos.cod_pedido) as varQuantPrazoTotal " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.tipo_pagamento = 'À Prazo') and (pedidos.TIPO_PEDIDO = 'VENDA') and pedidos.cancelado = 0 and (pedidos.codcaixa = " & Caixa_Controle_semOS.StatusBar1.Panels(3).Text & ") AND  pedidos.caixa = '" & StatusBar1.Panels(2).Text & "'  AND (parcelas.STATUS = 0)"
    
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vVlrVendasPrazoTotal = r("varSomaPrazoTotais") Else vVlrVendasPrazoTotal = 0
    
    'VENDAS CANCELADAS
    Dim vVlrVendasCanceladoTotal As Currency
    Dim vQtdeVendasCanceladoTotal As Integer
    sSQL = "SELECT ISNULL(SUM(TOTAL), 0) AS varSomaCanceladoTotais, count(cod_pedido) as varQuantCanceladoTotal " & _
           "FROM pedidos  " & _
           "WHERE pedidos.cancelado = 1 AND (codcaixa = " & txtCodCaixa.Text & ") " & Maquina_Venda
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vVlrVendasCanceladoTotal = r("varSomaCanceladoTotais"): vQtdeVendasCanceladoTotal = r("varQuantCanceladoTotal") Else vVlrVendasCanceladoTotal = 0: vQtdeVendasCanceladoTotal = 0

   'ORÇAMENTO
    Dim vVlrVendasOrcamentoTotal As Currency
    Dim vQtdeVendasOrcamentoTotal As Integer
    sSQL = "SELECT ISNULL(SUM(TOTAL), 0) AS varSomaOrcamentoTotais, count(cod_pedido) as varQuantOrcamentoTotal " & _
           "FROM pedidos  " & _
           "WHERE pedidos.cancelado = 1 AND (codcaixa = " & txtCodCaixa.Text & ") " & Maquina_Venda
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vVlrVendasOrcamentoTotal = r("varSomaOrcamentoTotais"): vQtdeVendasOrcamentoTotal = r("varQuantOrcamentoTotal") Else vVlrVendasOrcamentoTotal = 0: vQtdeVendasOrcamentoTotal = 0

   'CONSIGNADO
    Dim vVlrVendasConsignadoTotal As Currency
    Dim vQtdeVendasConsignadoTotal As Integer
    sSQL = "SELECT ISNULL(SUM(TOTAL), 0) AS varSomaConsignadoTotais, count(cod_pedido) as varQuantConsignadoTotal " & _
           "FROM pedidos  " & _
           "WHERE TIPO_PEDIDO = 'CONSIGNADO' AND pedidos.cancelado = 0 AND (codcaixa = " & txtCodCaixa.Text & ") " & Maquina_Venda
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vVlrVendasConsignadoTotal = r("varSomaConsignadoTotais"): vQtdeVendasConsignadoTotal = r("varQuantConsignadoTotal") Else vVlrVendasConsignadoTotal = 0: vQtdeVendasConsignadoTotal = 0
    
    'ALUGUEL
    Dim vVlrAluguelTotal As Currency
    Dim vQtdeAluguelTotal As Integer
    sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaTotaisAluguel, count(pedidos.COD_PEDIDO) as varQuantAluguelTotal " & _
           "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
           "WHERE (pedidos.codcaixa = " & txtCodCaixa.Text & ") AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'ALUGUEL') and pedidos.cancelado = 0 AND (parcelas.STATUS = 0) " & Maquina_Venda
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vVlrAluguelTotal = r("varSomaTotaisAluguel"): vQtdeAluguelTotal = r("varQuantAluguelTotal") Else vVlrAluguelTotal = 0: vQtdeAluguelTotal = 0

    'OS
    Dim vVlrOSTotal As Currency
    Dim vQtdeOSTotal As Integer
    sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaTotaisOS, count(pedidos.COD_PEDIDO) as varQuantOSTotal " & _
           "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
           "WHERE (pedidos.codcaixa = " & txtCodCaixa.Text & ") AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'OS') and pedidos.cancelado = 0 AND (parcelas.STATUS = 0) " & Maquina_Venda
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vVlrOSTotal = r("varSomaTotaisOS"): vQtdeOSTotal = r("varQuantOSTotal") Else vVlrOSTotal = 0: vQtdeOSTotal = 0
    
 

End If


'Mostrar_APrazo
'Mostrar_Retiradas

'FormatarGridResumido r
  
'If r.State <> 0 Then r.Close
'Set r = Nothing

'mostrar todas as saídas na folha
SQL = "SELECT HORA as vSHora, SUBDESCRICAO + '/' + DESCRICAO as vSDescricao, COD_FUNCIONARIO as vSFunc, VALOR as vSValor FROM caixa_saida " & _
       "WHERE FONTE = 'CAIXA ATUAL' AND (codcaixa = " & Caixa_Controle_semOS.StatusBar1.Panels(3).Text & ") AND (caixa = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "')"
Set r = dbData.OpenRecordset(SQL)

If r.EOF Then
    Set r = dbData.OpenRecordset(sSQL)
End If

Me.Hide
Dim vStrCaixa As String
Dim vStrCodCaixa As String
vStrCaixa = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
vStrCodCaixa = Caixa_Controle_semOS.StatusBar1.Panels(3).Text

If vAluguelAtiva = False And vOSAtiva = False Then
    Set REL_Caixa_Fech_Resumido.ReportMain1.Recordset = r
    REL_Caixa_Fech_Resumido.MostrarRetiradas vStrCaixa, vStrCodCaixa
    
    'Troco
    REL_Caixa_Fech_Resumido.rfExtraTroco.Caption = Format(vVlrTroco, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfExtraTrocoQtde.Caption = Format(1, "000") & " "
    
    'vendas
    REL_Caixa_Fech_Resumido.rfVendasTotal.Caption = Format(vVlrVendasTotal, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfVendasDinheiro.Caption = Format(vVlrVendasDinheiro, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfVendasPix.Caption = Format(vVlrVendasPix, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfVendasTransferencia.Caption = Format(vVlrVendasTransferencia, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfVendasDeposito.Caption = Format(vVlrVendasDeposito, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfVendasCartao.Caption = Format(vVlrVendasCartao, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfVendasCheque.Caption = Format(vVlrVendasCheque, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfVendasFinanceira.Caption = Format(vVlrVendasFinanceira, ocMONEY) & " "
    
    REL_Caixa_Fech_Resumido.rfVendasQtde.Caption = Format(vQtdeVendasTotal, "000") & " "
    REL_Caixa_Fech_Resumido.rfVendasDinheiroQtde.Caption = Format(vQtdeVendasDinheiro, "000") & " "
    REL_Caixa_Fech_Resumido.rfVendasPixQtde.Caption = Format(vQtdeVendasPix, "000") & " "
    REL_Caixa_Fech_Resumido.rfVendasTransferenciaQtde.Caption = Format(vQtdeVendasTransferencia, "000") & " "
    REL_Caixa_Fech_Resumido.rfVendasDepositoQtde.Caption = Format(vQtdeVendasDeposito, "000") & " "
    REL_Caixa_Fech_Resumido.rfVendasFinanceiraQtde.Caption = Format(vQtdeVendasFinanceira, "000") & " "
    REL_Caixa_Fech_Resumido.rfVendasCartaoQtde.Caption = Format(vQtdeVendasCartao, "000") & " "
    REL_Caixa_Fech_Resumido.rfVendasChequeQtde.Caption = Format(vQtdeVendasCheque, "000") & " "
    
    'parcelas
    REL_Caixa_Fech_Resumido.rfParcelasTotal.Caption = Format(vVlrParcelasTotal, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfParcelasDinheiro.Caption = Format(vVlrParcelasDinheiro, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfParcelasPix.Caption = Format(vVlrParcelasPix, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfParcelasTransferencia.Caption = Format(vVlrParcelasTransferencia, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfParcelasDeposito.Caption = Format(vVlrParcelasDeposito, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfParcelasCartao.Caption = Format(vVlrParcelasCartao, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfParcelasCheque.Caption = Format(vVlrParcelasCheque, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfParcelasFinanceira.Caption = Format(vVlrParcelasFinanceira, ocMONEY) & " "
    
    REL_Caixa_Fech_Resumido.rfParcelasTotalQtde.Caption = Format(vQtdeParcelasTotal, "000") & " "
    REL_Caixa_Fech_Resumido.rfParcelasDinheiroQtde.Caption = Format(vQtdeParcelasDinheiro, "000") & " "
    REL_Caixa_Fech_Resumido.rfParcelasPixQtde.Caption = Format(vQtdeParcelasPix, "000") & " "
    REL_Caixa_Fech_Resumido.rfParcelasTransferenciaQtde.Caption = Format(vQtdeParcelasTransferencia, "000") & " "
    REL_Caixa_Fech_Resumido.rfParcelasDepositoQtde.Caption = Format(vQtdeParcelasDeposito, "000") & " "
    REL_Caixa_Fech_Resumido.rfParcelasFinanceiraQtde.Caption = Format(vQtdeParcelasFinanceira, "000") & " "
    REL_Caixa_Fech_Resumido.rfParcelasCartaoQtde.Caption = Format(vQtdeParcelasCartao, "000") & " "
    REL_Caixa_Fech_Resumido.rfParcelasChequeQtde.Caption = Format(vQtdeParcelasCheque, "000") & " "
    
    'haveres
    REL_Caixa_Fech_Resumido.rfHaveresTotal.Caption = Format(vVlrHaveresTotal, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfHaveresDinheiro.Caption = Format(vVlrHaveresDinheiro, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfHaveresPix.Caption = Format(vVlrHaveresPix, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfHaveresTransferencia.Caption = Format(vVlrHaveresTransferencia, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfHaveresDeposito.Caption = Format(vVlrHaveresDeposito, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfHaveresCartao.Caption = Format(vVlrHaveresCartao, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfHaveresCheque.Caption = Format(vVlrHaveresCheque, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfHaveresFinanceira.Caption = Format(vVlrHaveresFinanceira, ocMONEY) & " "
    
    REL_Caixa_Fech_Resumido.rfHaveresTotalQtde.Caption = Format(vQtdeHaveresTotal, "000") & " "
    REL_Caixa_Fech_Resumido.rfHaveresDinheiroQtde.Caption = Format(vQtdeHaveresDinheiro, "000") & " "
    REL_Caixa_Fech_Resumido.rfHaveresPixQtde.Caption = Format(vQtdeHaveresPix, "000") & " "
    REL_Caixa_Fech_Resumido.rfHaveresTransferenciaQtde.Caption = Format(vQtdeHaveresTransferencia, "000") & " "
    REL_Caixa_Fech_Resumido.rfHaveresDepositoQtde.Caption = Format(vQtdeHaveresDeposito, "000") & " "
    REL_Caixa_Fech_Resumido.rfHaveresFinanceiraQtde.Caption = Format(vQtdeHaveresFinanceira, "000") & " "
    REL_Caixa_Fech_Resumido.rfHaveresCartaoQtde.Caption = Format(vQtdeHaveresCartao, "000") & " "
    REL_Caixa_Fech_Resumido.rfHaveresChequeQtde.Caption = Format(vQtdeHaveresCheque, "000") & " "
    
    'resumo
    Dim vResumoDinheiro As Currency
    Dim vResumoPix As Currency
    Dim vResumoTransferencia As Currency
    Dim vResumoDeposito As Currency
    Dim vResumoCartao As Currency
    Dim vResumoCheque As Currency
    Dim vResumoFinanceira As Currency
    
    vResumoDinheiro = vVlrVendasDinheiro + vVlrParcelasDinheiro + vVlrHaveresDinheiro + vVlrSuprimento
    vResumoPix = vVlrVendasPix + vVlrParcelasPix + vVlrHaveresPix
    vResumoTransferencia = vVlrVendasTransferencia + vVlrParcelasTransferencia + vVlrHaveresTransferencia
    vResumoDeposito = vVlrVendasDeposito + vVlrParcelasDeposito + vVlrHaveresDeposito
    vResumoCartao = vVlrVendasCartao + vVlrParcelasCartao + vVlrHaveresCartao
    vResumoCheque = vVlrVendasCheque + vVlrParcelasCheque + vVlrHaveresCheque
    vResumoFinanceira = vVlrVendasFinanceira + vVlrParcelasFinanceira + vVlrHaveresFinanceira
    
    REL_Caixa_Fech_Resumido.rfResumoDinheiro.Caption = Format(vResumoDinheiro, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfResumoPix.Caption = Format(vResumoPix, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfResumoTransferencia.Caption = Format(vResumoTransferencia, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfResumoDeposito.Caption = Format(vResumoDeposito, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfResumoCartao.Caption = Format(vResumoCartao, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfResumoCheque.Caption = Format(vResumoCheque, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfResumoFinanceira.Caption = Format(vResumoFinanceira, ocMONEY) & " "
    
    Dim vResumoDinheiroQtde As Integer
    Dim vResumoPixQtde As Integer
    Dim vResumoTransferenciaQtde As Integer
    Dim vResumoDepositoQtde As Integer
    Dim vResumoCartaoQtde As Integer
    Dim vResumoChequeQtde As Integer
    Dim vResumoFinanceiraQtde As Integer
    
    vResumoDinheiroQtde = vQtdeVendasDinheiro + vQtdeParcelasDinheiro + vQtdeHaveresDinheiro + vQtdeSuprimento
    vResumoPixQtde = vQtdeVendasPix + vQtdeParcelasPix + vQtdeHaveresPix
    vResumoTransferenciaQtde = vQtdeVendasTransferencia + vQtdeParcelasTransferencia + vQtdeHaveresTransferencia
    vResumoDepositoQtde = vQtdeVendasDeposito + vQtdeParcelasDeposito + vQtdeHaveresDeposito
    vResumoFinanceiraQtde = vQtdeVendasFinanceira + vQtdeParcelasFinanceira + vQtdeHaveresFinanceira
    vResumoCartaoQtde = vQtdeVendasCartao + vQtdeParcelasCartao + vQtdeHaveresCartao
    vResumoChequeQtde = vQtdeVendasCheque + vQtdeParcelasCheque + vQtdeHaveresCheque
    
    REL_Caixa_Fech_Resumido.rfResumoDinheiroQtde.Caption = Format(vResumoDinheiroQtde, "000") & " "
    REL_Caixa_Fech_Resumido.rfResumoPixQtde.Caption = Format(vResumoPixQtde, "000") & " "
    REL_Caixa_Fech_Resumido.rfResumoTransferenciaQtde.Caption = Format(vResumoTransferenciaQtde, "000") & " "
    REL_Caixa_Fech_Resumido.rfResumoDepositoQtde.Caption = Format(vResumoDepositoQtde, "000") & " "
    REL_Caixa_Fech_Resumido.rfResumoFinanceiraQtde.Caption = Format(vResumoFinanceiraQtde, "000") & " "
    REL_Caixa_Fech_Resumido.rfResumoCartaoQtde.Caption = Format(vResumoCartaoQtde, "000") & " "
    REL_Caixa_Fech_Resumido.rfResumoChequeQtde.Caption = Format(vResumoChequeQtde, "000") & " "
    
    'suprimentos
    REL_Caixa_Fech_Resumido.rfSuprimentoTotal.Caption = Format(vVlrSuprimento, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfSuprimentoDinheiro.Caption = Format(vVlrSuprimento, ocMONEY) & " "
    
    REL_Caixa_Fech_Resumido.rfSuprimentoTotalQtde.Caption = Format(vQtdeSuprimento, "000") & " "
    REL_Caixa_Fech_Resumido.rfSuprimentoDinheiroQtde.Caption = Format(vQtdeSuprimento, "000") & " "
    
    'sangrias
    REL_Caixa_Fech_Resumido.rfSangriaTotal.Caption = Format(vVlrSangria, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfSangriaDinheiro.Caption = Format(vVlrSangria, ocMONEY) & " "
    
    REL_Caixa_Fech_Resumido.rfSangriaTotalQtde.Caption = Format(vQtdeSangria, "000") & " "
    REL_Caixa_Fech_Resumido.rfSangriaDinheiroQtde.Caption = Format(vQtdeSangria, "000") & " "
    
    'sangrias
    REL_Caixa_Fech_Resumido.rfSangriaTotal.Caption = Format(vVlrSangria, ocMONEY) & " "
    REL_Caixa_Fech_Resumido.rfSangriaDinheiro.Caption = Format(vVlrSangria, ocMONEY) & " "
    
    REL_Caixa_Fech_Resumido.rfSangriaTotalQtde.Caption = Format(vQtdeSangria, "000") & " "
    REL_Caixa_Fech_Resumido.rfSangriaDinheiroQtde.Caption = Format(vQtdeSangria, "000") & " "
    
    'extra
    REL_Caixa_Fech_Resumido.rfeExtraAPrazoQtde.Caption = Format(vQtdeVendasPrazoTotal, "000") & " "
    REL_Caixa_Fech_Resumido.rfeExtraAPrazo.Caption = Format(vVlrVendasPrazoTotal, ocMONEY) & " "
    
    REL_Caixa_Fech_Resumido.rfeExtraCanceladasQtde.Caption = Format(vQtdeVendasCanceladoTotal, "000") & " "
    REL_Caixa_Fech_Resumido.rfeExtraCanceladas.Caption = Format(vVlrVendasCanceladoTotal, ocMONEY) & " "
    
    REL_Caixa_Fech_Resumido.rfeExtraOrcamentosQtde.Caption = Format(vQtdeVendasOrcamentoTotal, "000") & " "
    REL_Caixa_Fech_Resumido.rfeExtraOrcamentos.Caption = Format(vVlrVendasOrcamentoTotal, ocMONEY) & " "
    
    REL_Caixa_Fech_Resumido.rfeExtraConsignadoQtde.Caption = Format(vQtdeVendasConsignadoTotal, "000") & " "
    REL_Caixa_Fech_Resumido.rfeExtraConsignado.Caption = Format(vVlrVendasConsignadoTotal, ocMONEY) & " "
    
    REL_Caixa_Fech_Resumido.rfeExtraAluguelQtde.Caption = Format(vQtdeAluguelTotal, "000") & " "
    REL_Caixa_Fech_Resumido.rfeExtraAluguel.Caption = Format(vVlrAluguelTotal, ocMONEY) & " "
    
    REL_Caixa_Fech_Resumido.rfeExtraOSQtde.Caption = Format(vVlrOSTotal, "000") & " "
    REL_Caixa_Fech_Resumido.rfeExtraOS.Caption = Format(vVlrOSTotal, ocMONEY) & " "
    
    'Saldos
    Dim vSaldoGeral As Currency
    vSaldoGeral = vVlrVendasTotal + vVlrParcelasTotal + vVlrHaveresTotal + vVlrSuprimento
    vSaldoGeral = vSaldoGeral - vVlrSangria
    REL_Caixa_Fech_Resumido.rfSaldoGeral.Caption = Format(vSaldoGeral, ocMONEY) & " "
    
    Dim vSaldoFisico As Currency
    vSaldoFisico = vVlrVendasDinheiro + vVlrParcelasDinheiro + vVlrHaveresDinheiro + vVlrSuprimento
    vSaldoFisico = vSaldoFisico + vVlrVendasCheque + vVlrParcelasCheque + vVlrHaveresCheque
    vSaldoFisico = vSaldoFisico - vVlrSangria
    REL_Caixa_Fech_Resumido.rfSaldoFisico.Caption = Format(vSaldoFisico, ocMONEY) & " "
    
    
    '===========================RODAPÉ
    Dim sSQLusuario As String
    Dim r_usuario As ADODB.Recordset
    
    sSQLusuario = "SELECT DATA_ABERTURA, HORA_ABERTURA, COD_FUNC_ABERTURA, DATA_FECHAMENTO, HORA_FECHAMENTO, COD_FUNC_FECHAMENTO, (CASE WHEN status = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS VarStatus, " & _
            "(SELECT Usuario.Login FROM Usuario INNER JOIN caixa_dia ON Usuario.Codigo = caixa_dia.COD_FUNC_ABERTURA wHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ")) AS Nome_Abertura, " & _
            "(SELECT Usuario_2.Login FROM Usuario AS Usuario_2 INNER JOIN caixa_dia AS caixa_dia_2 ON Usuario_2.Codigo = caixa_dia_2.COD_FUNC_FECHAMENTO WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ")) AS Nome_Fechamento " & _
           "FROM caixa_dia AS caixa_dia_1 " & _
           "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & txtCodCaixa.Text & ");"
    Set r_usuario = dbData.OpenRecordset(sSQLusuario)
    
    
    If Not r_usuario.EOF Then
        REL_Caixa_Fech_Resumido.rfCodUsuarioA.Caption = Format(r_usuario("COD_FUNC_ABERTURA"), "00")
        REL_Caixa_Fech_Resumido.rfNomeUsuarioA.Caption = ValidateNull(r_usuario("Nome_Abertura"))
        REL_Caixa_Fech_Resumido.rfDataA.Caption = Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
        REL_Caixa_Fech_Resumido.rfHoraA.Caption = Format(ValidateNull(r_usuario("HORA_ABERTURA")), "hh:mm")
        
        REL_Caixa_Fech_Resumido.rfNomeUsuarioF.Caption = ValidateNull(r_usuario("Nome_Fechamento"))
        If IsNull(r_usuario("DATA_FECHAMENTO")) Then
            REL_Caixa_Fech_Resumido.rfDataF.Caption = ""
            REL_Caixa_Fech_Resumido.rfCodUsuarioF.Caption = ""
            REL_Caixa_Fech_Resumido.rfHoraF.Caption = ""
        Else
            REL_Caixa_Fech_Resumido.rfCodUsuarioF.Caption = Format(ValidateNull(r_usuario("COD_FUNC_FECHAMENTO")), "00")
            REL_Caixa_Fech_Resumido.rfDataF.Caption = Format(ValidateNull(r_usuario("DATA_FECHAMENTO")), "dd/mm/yyyy")
            REL_Caixa_Fech_Resumido.rfHoraF.Caption = Format(ValidateNull(r_usuario("HORA_FECHAMENTO")), "hh:mm")
        End If
    
        REL_Caixa_Fech_Resumido.rfSituacao.Caption = ValidateNull(r_usuario("VARSTATUS"))
    End If
    
    REL_Caixa_Fech_Resumido.rfCaixa.Caption = StatusBar1.Panels(2).Text
    REL_Caixa_Fech_Resumido.rfCodCaixa.Caption = Format(txtCodCaixa.Text, "0000")
    
    
    REL_Caixa_Fech_Resumido.ReportMain1.NomeImpressora = var_ImpNormal
    REL_Caixa_Fech_Resumido.ReportMain1.Ativar
End If
Me.Show
End Sub

Private Sub InutilizarCuponsFiscais()
If vConfImprimeNFCeLocal = "SIM" Then
    Dim NFCeContingencia As Boolean
    Dim codPedido As String, nNota As String, CNPJ As String
    Dim IdNFProd As Long
    
    sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
    Set r = dbData.OpenRecordset(sSQL)
    NFCeContingencia = r!ContigenciaNFCe
    
    If NFCeContingencia Then
       MsgBox "CONTINGÊNCIA DA NFCE ATIVADA, INUTILIZAÇÃO NÃO PERMITIDA!", vbInformation + vbOKOnly
       GoTo Caifora
    End If
    
    CNPJ = SQLExecutaRetorno("SELECT CNPJ FROM Empresa", "CNPJ", "")
    
    sSQL = "SELECT IdNFProd, NumeNota " & _
           "FROM TbNFCe " & _
           "WHERE (DataEmissao = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103)) AND (NFCeEnviada = 0) AND (NFCeCancelada = 0) AND (Inutilizada = 0)"
    Set r = dbData.OpenRecordset(sSQL)
    
    'Debug.Print sSQL
    
    Dim sistNFe As snfe.Util
    Set sistNFe = New snfe.Util
    xCaminhoXML = ""
    
    'frameAguarde.Visible = True
    DoEvents
    Do While Not r.EOF
        iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFe)
        nNota = r!NumeNota
        iRetorno = sistNFe.InutilizarNumeracao(Format(Date, "yyyy"), CNPJ, "ERRO AO TRANSMITIR NOTA, PERDA DE SEQUENCIA", nNota, nNota, 1, xCaminhoXML)
        cStat = sistNFe.retInutilizacao.infInut.cStat
        NFeMotivo = sistNFe.retInutilizacao.infInut.xMotivo
        NFeDataHora = sistNFe.retInutilizacao.infInut.dhRecbto
        NFeNumeroProtocolo = sistNFe.retInutilizacao.infInut.nProt
        If cStat = 102 Or cStat = 563 Then
           sSQL = "UPDATE TbNFCe SET Inutilizada = 1, NFCeProtocolo = " & NFeNumeroProtocolo & ", NFCeProtocoloDataHora = '" & NFeDataHora & "', Num_OS_VD_Origem = 0 WHERE IdNFProd = " & r!IdNFProd
           vgDb.Execute sSQL
        Else
        End If
    
        r.MoveNext
    Loop
    
End If

Caifora:
    Set r = Nothing
    'frameAguarde.Visible = False
    DoEvents
    Screen.MousePointer = vbDefault
    Set sistNFe = Nothing
End Sub

Private Sub SaberSaldoAnteriorUpdate()
sSQL = "SELECT ISNULL(saldo_atual, 0) AS varUltimoSaldoAtual FROM caixa_saldo where (codigo = " & varCodSaldo & " - 1);"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then ANTERIOR = r("varUltimoSaldoAtual")

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub SaberSaldoAnterior()
sSQL = "SELECT TOP 1 ISNULL(saldo_atual, 0) AS varUltimoSaldoAtual FROM caixa_saldo ORDER BY codigo DESC;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then ANTERIOR = r("varUltimoSaldoAtual")

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub



Private Sub cmdReativar_Click()
Dim var_CodSequencia As Long
Dim var_CodCaixa As Long

If txtCodFuncAP.Text = "" Then MsgBox "Coloque o Cód de Funcinario!", vbInformation, "Aviso do Sistema": txtCodFuncAP.SetFocus: Exit Sub

If cmdReativar.Caption = "Abrir Caixa" Then
   'sequencia de caixas numeração
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_caixa FROM caixa_dia"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then var_CodSequencia = r("ultimo_caixa") + 1
   
   'sequencia de caixas numeração
   sSQL = "SELECT ISNULL(MAX(CODCAIXA), 0) AS ultimo_caixa FROM caixa_dia where CAIXA = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "'"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then var_CodCaixa = r("ultimo_caixa") + 1
   
   dbData.Execute "INSERT INTO caixa_dia (codigo, CODCAIXA, data_abertura, hora_abertura, COD_FUNC_ABERTURA, status, entrada, saida, saldo, caixa) VALUES (" & var_CodSequencia & ", " & var_CodCaixa & ", CONVERT(DATETIME, '" & Format(StatusBar1.Panels(4).Text, ocDATA) & "', 103), '" & Format(Now, ocHRMN) & "', " & txtCodFuncAP.Text & ", 0, " & Replace(CCur(lblEntrada), ",", ".") & ", " & Replace(CCur(lblSaida), ",", ".") & ", " & Replace(CCur(lblTotal), ",", ".") & ", '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "');"


    'caixa troco
    Dim x_Troco As Long
    
    If txtTroco.Text = "" Then Exit Sub
    
    x_Troco = 1
    sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_troco FROM caixa_troco where (caixa = '" & StatusBar1.Panels(2).Text & "');"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then x_Troco = r("ultimo_troco") + 1
    If r.State <> 0 Then r.Close
    Set r = Nothing
    
    dbData.Execute "INSERT INTO caixa_troco (codigo, data, valor, caixa, codcaixa) VALUES (" & x_Troco & ", CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103), " & Replace(CCur(txtTroco.Text), ",", ".") & ", '" & StatusBar1.Panels(2).Text & "', " & var_CodCaixa & ");"

Else
   dbData.Execute "UPDATE caixa_dia SET status = 0 WHERE (data_abertura = CONVERT(DATETIME, '" & Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, ocDATA) & "', 103)) AND (caixa = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "');"
   ''execSQL "DELETE FROM CAIXA_DIA WHERE DATA_ABERTURA = #" & Format(StatusBar1.Panels(3).Text, "dd/mm/yyyy") & "# and CAIXA = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "' "
   'dbData.Execute "DELETE FROM caixa_saldo WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103));"
End If

sSQL = "SELECT VencimentoCert FROM empresa"
Set r = dbData.OpenRecordset(sSQL)

If Not IsNull(r!VencimentoCert) Then
   If (DateDiff("d", Date, r!VencimentoCert) <= 10) Then
      If (DateDiff("d", Date, r!VencimentoCert) <= 0) Then
         MsgBox "Seu certificado digital venceu." & vbNewLine & "Todo ano tem que renovar." & vbNewLine & "Mande mensagem para seu contador para ele renovar para você. Assim que ele te enviar o arquivo, instalamos para você!", vbCritical + vbOKOnly, "CERTIFICADO DIGITAL VENCIDO"
      ElseIf (DateDiff("d", Date, r!VencimentoCert) = 1) Then
         MsgBox "Seu certificado digital vence hoje!" & vbNewLine & "Todo ano tem que renovar." & vbNewLine & "Mande mensagem para seu contador para ele renovar para você. Assim que ele te enviar o arquivo, instalamos para você!", vbExclamation + vbOKOnly, "CERTIFICADO DIGITAL"
      Else
         MsgBox "Faltam " & CStr(Int(DateDiff("d", r!VencimentoCert, Date)) * (-1)) & " dias para vencer o seu CERTIFICADO DIGITAL!", vbInformation + vbOKOnly, "CERTIFICADO DIGITAL"
      End If
   End If
End If
MsgBox "SEU CAIXA FOI ABERTO COM SUCESSO!!" & Chr(13) & "Reabra seu PDV novamente", vbInformation, "Aviso do Sistema"

If varFluxoCaixa = False Then
    KillApp "PDV.exe"
    'End
End If

varFluxoCaixa = False
Unload Me
'Unload PDV
End Sub

Private Sub cmdFecharCaixa_Click()
On Error GoTo TrataErro

If txtFuncAP.Text = "" Then
   ShowMsg "Faltou o código do funcionário!", vbExclamation
   txtCodFuncAP.SetFocus
   Exit Sub
End If

If lblEntrada.Caption = "" Or lblSaida.Caption = "" Or lblTotal.Caption = "" Then
   Exit Sub
End If

'SALVAR NA TABELA CAIXA_DIA
sSQL = "UPDATE caixa_dia SET " & _
   "entrada = " & Replace(CCur(lblEntrada), ",", ".") & ", " & _
   "saida = " & Replace(CCur(lblSaida), ",", ".") & ", " & _
   "saldo = " & Replace(CCur(lblTotal), ",", ".") & ", " & _
   "status = 1, " & _
   "data_fechamento = CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103), " & _
   "hora_fechamento = '" & Format(Now, ocHRMN) & "', " & _
   "COD_FUNC_FECHAMENTO = " & txtCodFuncAP & _
   " WHERE (codcaixa = " & txtCodCaixa.Text & ") and (caixa = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "');"
'Debug.Print sSQL
dbData.Execute sSQL

GerarSaldo
'InutilizarCuponsFiscais

'verificar se a maquina é o servidor
Dim vNomeMaquina As String
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
vNomeMaquina = oIni.LerTexto("IP_MAQUINA", "ip")
Set oIni = Nothing

'se for servidor faz o backup
If Left(vNomeMaquina, 1) = "." Then
    Backup
End If

'If ShowMsg("Deseja imprimir o caixa: '" & StatusBar1.Panels(2).Text & "' com Cód.Caixa No. '" & txtCodCaixa.Text & "' ??", vbInformation + vbYesNo) = vbYes Then
'    ImprimirCaixaResumido
'End If

'If ShowMsg("Deseja imprimir um relatório com todos os produtos vendido hoje?", vbInformation + vbYesNo) = vbYes Then
'    'Load Vendas_Consulta_PorProdutosAgrupadas
'    vOrigemRelatorio = True
'    Vendas_Consulta_PorProdutosAgrupadas.cboCriterioPrinc.Text = "DATA"
'    Vendas_Consulta_PorProdutosAgrupadas.lblFim.Visible = False
'    Vendas_Consulta_PorProdutosAgrupadas.mskFim.Visible = False
'    Vendas_Consulta_PorProdutosAgrupadas.lblAte.Visible = False
'    Vendas_Consulta_PorProdutosAgrupadas.cmdCalendario1.Visible = True
'    Vendas_Consulta_PorProdutosAgrupadas.cmdCalendario2.Visible = False
'    Vendas_Consulta_PorProdutosAgrupadas.lblMes.Visible = False
'    Vendas_Consulta_PorProdutosAgrupadas.cboMes.Visible = False
'    Vendas_Consulta_PorProdutosAgrupadas.lblAno.Visible = False
'    Vendas_Consulta_PorProdutosAgrupadas.cboAno.Visible = False
'    Vendas_Consulta_PorProdutosAgrupadas.lblInicio.Visible = True
'    Vendas_Consulta_PorProdutosAgrupadas.lblInicio.Caption = "Data"
'    Vendas_Consulta_PorProdutosAgrupadas.mskInicio.Visible = True
'    Vendas_Consulta_PorProdutosAgrupadas.mskInicio.Text = Format(Date, "dd/mm/yyyy")
'    Vendas_Consulta_PorProdutosAgrupadas.cmdLocalizar_Click
'    Vendas_Consulta_PorProdutosAgrupadas.cmdImprimir_Click
'    vOrigemRelatorio = False
'End If

'If ShowMsg("Deseja imprimir um relatório com todos os produtos vendido hoje?", vbInformation + vbYesNo) = vbYes Then
'End If

If varFluxoCaixa = False Then
    KillApp "PDV.exe"
    'End
End If

Unload Me
Unload Caixa_Controle_semOS

varFluxoCaixa = False
Exit Sub

TrataErro:
If Err.Number = 3022 Then
   ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
   Exit Sub
End If
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


ANTERIOR = 0
Dim var_Entrada As Currency, var_Saida As Currency, var_Total As Currency
Dim sSQL As String
Dim r As ADODB.Recordset

Dim var_Caixa As String     'colocar o nome da maquina na barra de status
'Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Caixa = oIni.LerTexto("DADOS_CAIXA", "caixa")
'Set oIni = Nothing

vConfImprimeNFCeLocal = oIni.LerTexto("IMPRIMIR_NFCE", "resposta")
Set oIni = Nothing

'mostrar os objetos de OS e/ou Aluguel
Dim oCfg As ConfigItem
Dim bStatus As Boolean

'os
Set oCfg = sysConfig("os")    'Recupera a config deseja
bStatus = CBool(oCfg.Value)   'Converte o valor para booleano
vOSAtiva = CBool(oCfg.Value)
Set oCfg = Nothing            'Destroi o objeto

'txtQuantDinheiroOS.Visible = bStatus 'Habilita/desabilida conforme valor
'txtTotalDinheiroOS.Visible = bStatus
'lblOS.Visible = bStatus

'aluguel
Set oCfg = sysConfig("aluguel")    'Recupera a config deseja
bStatus = CBool(oCfg.Value)   'Converte o valor para booleano
vAluguelAtiva = CBool(oCfg.Value)
'Set oCfg = Nothing            'Destroi o objeto

   
StatusBar1.Panels(2).Text = var_Caixa
'StatusBar1.Panels(2).Text = "CAIXA01"  'ESSE ==================


'txtCodCaixa.Text = Caixa_Controle_semOS.txtCodCaixa.Text 'ver depois
 StatusBar1.Panels(4).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
 mskData.Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
 txtHora.Text = Format(Now, "HH:MM")
 txtTroco.Text = Format(0, ocMONEY)

'mskData.Text = Format(Date, "dd/mm/yy")
If Caixa_Controle_semOS.txtEntrada.Text = "" Then lblEntrada.Caption = "0" Else lblEntrada.Caption = Caixa_Controle_semOS.txtEntrada.Text
If Caixa_Controle_semOS.txtSaida.Text = "" Then lblSaida.Caption = "0" Else lblSaida.Caption = Caixa_Controle_semOS.txtSaida.Text
'lblEntrada.Caption = "25"
'lblSaida.Caption = "5"
var_Entrada = lblEntrada.Caption  'ESSE ==================
var_Saida = lblSaida.Caption        'ESSE ==================

var_Total = var_Entrada - var_Saida
lblTotal.Caption = Format(var_Total, ocMONEY)

'MOSTRAR SE O CAIXA ESTÁ FECHADO
'sSQL = "SELECT codigo, codcaixa, data_abertura, caixa, status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, ocDATA) & "', 103)) AND (caixa = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "');"
sSQL = "SELECT codigo, codcaixa, data_abertura, caixa, status FROM caixa_dia WHERE (caixa = '" & Caixa_Controle_semOS.StatusBar1.Panels(2).Text & "') and status = 0;"
Set r = dbData.OpenRecordset(sSQL)

 If r.BOF Then
    cmdReativar.Caption = "Abrir Caixa"
    cmdReativar.Enabled = True       'tem que apagar essa linha depois
    cmdFecharCaixa.Enabled = False
 Else
    If CInt(ValidateNull(r("status"))) = 0 Then
       txtCodCaixa.Text = CInt(ValidateNull(r("codcaixa")))
       cmdReativar.Enabled = False
       cmdFecharCaixa.Enabled = True
    Else
       txtCodCaixa.Text = CInt(ValidateNull(r("codcaixa")))
       cmdReativar.Enabled = True
       cmdReativar.Caption = "Reativar Caixa"
       cmdFecharCaixa.Enabled = False
    End If
 End If
 
 'txtCodCaixa.Text = "139"   'ESSE ==================
 
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

Private Sub txtTroco_GotFocus()
SelectControl txtTroco
End Sub


Private Sub txtTroco_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtTroco_LostFocus()
If txtTroco.Text = "" Then txtTroco.Text = Format(0, ocMONEY)
txtTroco.Text = Format(txtTroco.Text, ocMONEY)
End Sub


