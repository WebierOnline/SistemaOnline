VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Exportar_XML 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXPORTAR XML"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6735
   Icon            =   "Exportar_XML2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   6585
      TabIndex        =   10
      Top             =   30
      Width           =   6615
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EXPORTAÇÃO XML"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   840
         TabIndex        =   11
         Top             =   180
         Width           =   2970
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   60
      TabIndex        =   1
      Top             =   900
      Width           =   6615
      Begin VB.CheckBox chkIncluirPDF 
         Caption         =   "Incluir PDF"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   1560
         Width           =   1155
      End
      Begin VB.CheckBox chkIncluirEntradas 
         Caption         =   "Incluir Entradas"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtDiretorioXML 
         Height          =   315
         Left            =   60
         TabIndex        =   0
         Top             =   1200
         Width           =   6195
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   540
         Width           =   1515
      End
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   1620
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   975
      End
      Begin ChamaleonBtn.chameleonButton chameleonButton1 
         Height          =   255
         Left            =   6300
         TabIndex        =   15
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "..."
         ENAB            =   -1  'True
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
         MICON           =   "Exportar_XML2.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblDirXML 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diretório de Destino:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label lblMes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   345
      End
      Begin VB.Label lblAno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         TabIndex        =   4
         Top             =   300
         Width           =   345
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   9
      Top             =   3435
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7541
            Text            =   "Online.Info - Informática"
            TextSave        =   "Online.Info - Informática"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "18:06"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin ChamaleonBtn.chameleonButton cmdEnviar 
      Height          =   435
      Left            =   1680
      TabIndex        =   12
      Top             =   2940
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Enviar"
      ENAB            =   -1  'True
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
      MICON           =   "Exportar_XML2.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCompactar 
      Height          =   435
      Left            =   60
      TabIndex        =   13
      Top             =   2940
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Compactar/Enviar"
      ENAB            =   -1  'True
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
      MICON           =   "Exportar_XML2.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblAguarde 
      AutoSize        =   -1  'True
      Caption         =   "Aguarde..."
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
      Height          =   300
      Left            =   3420
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "Exportar_XML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
  hwndOwner As Long
  pIDLRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type

Dim sSQL As String
Dim r As ADODB.Recordset
Dim rCont As ADODB.Recordset
Private moCombo As cComboHelper
Private Caminho As String
Dim vFantasia As String, vRazao As String, vCnpj As String, diretorioDestino As String, DiretorioOrigem As String
Dim vMes As String, vAno As String, vMesNum As String, vCaminhoXML

Private Sub cboAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = LastYear To FirstYear Step -1
   cboAno.AddItem i
Next
End Sub

Private Sub cboMes_GotFocus()
Dim vMes As Integer
cboMes.Clear
For vMes = 1 To 12
   cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
Next
moCombo.AttachTo cboMes
End Sub

Private Sub cmdCompactar_Click()
Dim FileZipName As String, PathToCompress As String, DestPath As String, FullPathZip As String
Dim NomeEmp As String, emailDestino As String, i As Integer, ComandoSQL As String
Dim rsEntradas As New ADODB.Recordset, rsNFe As New ADODB.Recordset, rsNFCe As New ADODB.Recordset
Dim xDiretorioDestino As String, xArquivoDestino As String

If cboMes.Text <> "" Then vMes = cboMes.Text Else: MsgBox "Escolha um mês!", vbInformation, "Aviso do Sistema": cboMes.SetFocus: Exit Sub
If cboAno.Text <> "" Then vAno = cboAno.Text Else: MsgBox "Escolha um ano!", vbInformation, "Aviso do Sistema": cboAno.SetFocus: Exit Sub

lblAguarde.Visible = True

If cboMes.Text = "Janeiro" Then
    vMesNum = 1
ElseIf cboMes.Text = "Fevereiro" Then
    vMesNum = 2
ElseIf cboMes.Text = "Março" Then
    vMesNum = 3
ElseIf cboMes.Text = "Abril" Then
    vMesNum = 4
ElseIf cboMes.Text = "Maio" Then
    vMesNum = 5
ElseIf cboMes.Text = "Junho" Then
    vMesNum = 6
ElseIf cboMes.Text = "Julho" Then
    vMesNum = 7
ElseIf cboMes.Text = "Agosto" Then
    vMesNum = 8
ElseIf cboMes.Text = "Setembro" Then
    vMesNum = 9
ElseIf cboMes.Text = "Outubro" Then
    vMesNum = 10
ElseIf cboMes.Text = "Novembro" Then
    vMesNum = 11
ElseIf cboMes.Text = "Dezembro" Then
    vMesNum = 12
End If

If Vazio(diretorioDestino) Then
  diretorioDestino = vCaminhoXML & "\ExportarXML"
  If Not Existe(diretorioDestino) Then MkDir diretorioDestino
  If Not Existe(diretorioDestino) Then MkDir GetDesktopFolder & "\" & "XML"
  'DiretorioDestino = DiretorioDestino & "\" & vCnpj
  'If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
  'DiretorioDestino = DiretorioDestino & "\" & vMes & vAno
  'If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
End If

'INICIAR COMPACTAÇÃO
If IniciaComponenteCompactacao Then

   i = 0
   
   'Caminho para comprimir arquivo
   DestPath = diretorioDestino
   
   'nome do arquivo
   NomeEmp = vRazao
   NomeEmp = RemoveAcento(NomeEmp)
   NomeEmp = Substitui(NomeEmp, ".,/", "", UM_A_UM)
   NomeEmp = Substitui(NomeEmp, " ", "_", UM_A_UM)
   FileZipName = NomeEmp & "_" & vMes & vAno & ".rar"
   
   'local de destino + ficheiro.rar
   diretorioDestino = vCaminhoXML & "\ExportarXML"
   FullPathZip = Transforma(diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & FileZipName)
   
   'Caminho a comprimir (nota : no final não deixar no fim da directoria o simbolo de '\'
   'DiretorioDestino = IIf(Right(DiretorioDestino, 1) = "\", Left(DiretorioDestino, Len(DiretorioDestino) - 1), DiretorioDestino)
   'C:\Sistemas\Projetos\Sistema\nfe\arquivos\procNFe\202107
   
   vMesNum = Format(vMesNum, "00")
   'DiretorioDestino = vCaminhoXML & "\nfe\arquivos\procNFe" & "\" & vAno & vMes
   DiretorioOrigem = vCaminhoXML & "\nfe\arquivos\procNFe" & "\" & vAno & vMesNum
   
   If chkIncluirPDF.Value = 1 Then
      ComandoSQL = "SELECT ChavedeAcesso " & _
                   "FROM NotaFiscal " & _
                   "WHERE MONTH(DataEmissao) = " & vMesNum & " AND YEAR(DataEmissao) = " & vAno
      Set rsNFe = dbData.OpenRecordset(ComandoSQL)
      Do While Not rsNFe.EOF
          xCaminhoXML = vCaminhoXML & "nfe\arquivos\PDF\NFe" & rsNFe!ChavedeAcesso & ".pdf"
          xDiretorioDestino = DiretorioOrigem & IIf(Right(DiretorioOrigem, 1) = "\", "", "\") & "PDF\"
          If Not Existe(xDiretorioDestino) Then MkDir xDiretorioDestino
          xArquivoDestino = xDiretorioDestino & rsNFe!ChavedeAcesso & ".pdf"
          If Existe(xCaminhoXML) = -1 Then FileCopy xCaminhoXML, xArquivoDestino
          rsNFe.MoveNext
      Loop
      DoEvents
      ComandoSQL = "SELECT NFCeChaveAcesso " & _
                   "FROM TbNFCe " & _
                   "WHERE MONTH(DataEmissao) = " & vMesNum & " AND YEAR(DataEmissao) = " & vAno
      Set rsNFCe = dbData.OpenRecordset(ComandoSQL)
      Do While Not rsNFCe.EOF
          xCaminhoXML = vCaminhoXML & "nfe\arquivos\PDF\NFe" & rsNFCe!NFCeChaveAcesso & ".pdf"
          xDiretorioDestino = DiretorioOrigem & IIf(Right(DiretorioOrigem, 1) = "\", "", "\") & "PDF\"
          If Not Existe(xDiretorioDestino) Then MkDir xDiretorioDestino
          xArquivoDestino = xDiretorioDestino & rsNFCe!NFCeChaveAcesso & ".pdf"
          If Existe(xCaminhoXML) = -1 Then FileCopy xCaminhoXML, xArquivoDestino
          rsNFCe.MoveNext
      Loop
   End If
   
   If chkIncluirEntradas.Value = 1 Then
      ComandoSQL = "SELECT ChavedeAcesso " & _
                   "FROM EntradaEstoque " & _
                   "WHERE MONTH(DataEmissao) = " & vMesNum & " AND YEAR(DataEmissao) = " & vAno
       Set rsEntradas = dbData.OpenRecordset(ComandoSQL)
       Do While Not rsEntradas.EOF
          xCaminhoXML = vCaminhoXML & "nfe\arquivos\ConfRecebto\" & rsEntradas!ChavedeAcesso & "-procNFe.xml"
          xDiretorioDestino = DiretorioOrigem & IIf(Right(DiretorioOrigem, 1) = "\", "", "\") & "Entradas\"
          If Not Existe(xDiretorioDestino) Then MkDir xDiretorioDestino
          xArquivoDestino = xDiretorioDestino & rsEntradas!ChavedeAcesso & "-procNFe.xml"
          If Existe(xCaminhoXML) = -1 Then FileCopy xCaminhoXML, xArquivoDestino
          rsEntradas.MoveNext
       Loop
   End If
   
   If Not Existe(DiretorioOrigem) Then MsgBox "Não existe a pasta referente ao mês selecionado!", vbInformation, "Aviso do Sistema": lblAguarde.Visible = False: Exit Sub
   
   PathToCompress = Transforma(DiretorioOrigem)
   
   'Chama o compressor que se encontra instalado para o efeito.
   If xWinRar <> "" Then
       Shell xWinRar & " a -ep1 " & FullPathZip & " " & PathToCompress, vbNormalFocus  ', vbHide
   Else
       Shell xWinZip & " -a -ep1 " & FullPathZip & " " & PathToCompress, vbNormalFocus ', vbHide
   End If
   
   DoEvents
   
   'Load frmECFMsg
   'frmECFMsg.SetaMensagem "Aguarde! Compactando arquivos XML..."
  
   'entra em loop até a criação do arquivo
   Do While Dir$(diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & FileZipName) = ""
       Sleep (1000)
   Loop
   
   Sleep 2000
   
   sSQL = "SELECT * FROM TbContabilista"
   Set rCont = dbData.OpenRecordset(sSQL)
   
   'If ckEnviarEmail Then
      Do While Not (Existe(diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & FileZipName) = -1)
         DoEvents
         i = i + 1
         If i > 50 Then Exit Do
      Loop
      If Not (Existe(diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & FileZipName) = -1) Then Exit Sub: lblAguarde.Visible = False
      If rCont.RecordCount > 0 Then emailDestino = rCont!email
      emailDestino = InputBox("Informe o Email do destinatario", "Envio de Email", emailDestino)
      If Not Vazio(emailDestino) Then
         'HourglassShow
         EnviaEmail emailDestino, FullPathZip
         DoEvents
         'HourglassClose
      End If
   'End If
   
   Set rCont = Nothing
   Set rsNFe = Nothing
   Set rsNFCe = Nothing
   Set rsEntradas = Nothing
End If

lblAguarde.Visible = False
End Sub

Private Sub EnviaEmail(EmailPara As String, Anexo1 As String)
Dim emailDest As String, pathAnexo() As String, NomeRemetente As String, corpoEmail As String, emailCC() As String
Dim temParcelas As Boolean
Dim sistNFe As snfe.Util

'On Error GoTo DeuErro

Set sistNFe = New snfe.Util

emailDest = EmailPara

ReDim emailCC(0)
emailCC(0) = EmailPara

If Vazio(emailDest) Then Exit Sub

'iRetorno = ConfiguraDLLNFeNFCe(IdFilial, 55, "1", sistNFe, False)
iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)

ReDim pathAnexo(0)
pathAnexo(0) = Anexo1

NomeRemetente = SQLExecutaRetorno("SELECT Fantasia FROM empresa", "Fantasia", r!Fantasia)
corpoEmail = "Segue em anexo os arquivos XML de NFe e NFCe emitidos no período. " & _
             "<br><br>" & _
             "Atenciosamente, " & _
             "<br><br>" & _
             "#nome_emitente#"
corpoEmail = Substitui(corpoEmail, "#nome_emitente#", SQLExecutaRetorno("SELECT RAZAO FROM empresa", "RAZAO"), SO_UM)

If (emailDest <> Empty) Then
   Screen.MousePointer = vbHourglass
   'GeraLogAcao LoadGasString(295) + " Enviando Email XML para Contabilidade"
   iRetorno = sistNFe.EmailEnviar(emailDest, "Arquivos XML ref. Mês " & LPad(vMesNum, 2, "0") & "/" & LPad(vAno, 4, "0"), corpoEmail, pathAnexo, emailCC)
   Screen.MousePointer = vbDefault
End If

If iRetorno Then MsgBox "Email enviado com sucesso!", vbInformation + vbOKOnly, "EMAIL OK!"

Set sistNFe = Nothing

Exit Sub
    
'DeuErro:
'    MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO: Envio Email"
'    Err.Clear
'    Set sistNFe = Nothing
End Sub

Private Sub cmdDiretorioXML_Click()
  'Opens a Treeview control that displays the directories in a computer
  Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
  
  szTitle = vbCr & vbCr & "Selecione a Pasta Desejada:"
  
  With tBrowseInfo
    .hwndOwner = Me.hwnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    txtDiretorioXML.Text = sBuffer
  End If
End Sub

Private Sub cmdEnviar_Click()
Dim FileZipName As String, PathToCompress As String, DestPath As String, FullPathZip As String
Dim NomeEmp As String, emailDestino As String, i As Integer, ComandoSQL As String

If cboMes.Text <> "" Then vMes = cboMes.Text Else: MsgBox "Escolha um mês!", vbInformation, "Aviso do Sistema": cboMes.SetFocus: Exit Sub
If cboAno.Text <> "" Then vAno = cboAno.Text Else: MsgBox "Escolha um ano!", vbInformation, "Aviso do Sistema": cboAno.SetFocus: Exit Sub

lblAguarde.Visible = True

   If cboMes.Text = "Janeiro" Then
      vMesNum = 1
   ElseIf cboMes.Text = "Fevereiro" Then
      vMesNum = 2
   ElseIf cboMes.Text = "Março" Then
      vMesNum = 3
   ElseIf cboMes.Text = "Abril" Then
      vMesNum = 4
   ElseIf cboMes.Text = "Maio" Then
      vMesNum = 5
   ElseIf cboMes.Text = "Junho" Then
      vMesNum = 6
   ElseIf cboMes.Text = "Julho" Then
      vMesNum = 7
   ElseIf cboMes.Text = "Agosto" Then
      vMesNum = 8
   ElseIf cboMes.Text = "Setembro" Then
      vMesNum = 9
   ElseIf cboMes.Text = "Outubro" Then
      vMesNum = 10
   ElseIf cboMes.Text = "Novembro" Then
      vMesNum = 11
   ElseIf cboMes.Text = "Dezembro" Then
      vMesNum = 12
   End If

   NomeEmp = vRazao
   NomeEmp = RemoveAcento(NomeEmp)
   NomeEmp = Substitui(NomeEmp, ".,/", "", UM_A_UM)
   NomeEmp = Substitui(NomeEmp, " ", "_", UM_A_UM)
   FileZipName = NomeEmp & "_" & vMes & vAno & ".rar"
   
   'local de destino + ficheiro.rar
   diretorioDestino = vCaminhoXML & "\ExportarXML"
   FullPathZip = Transforma(diretorioDestino & IIf(Right(diretorioDestino, 1) = "\", "", "\") & FileZipName)
   
   If Not Existe(FullPathZip) Then MsgBox "Não existe a pasta referente ao mês selecionado!", vbInformation, "Aviso do Sistema": Exit Sub

   sSQL = "SELECT * FROM TbContabilista"
   Set rCont = dbData.OpenRecordset(sSQL)
      
   If Not (Existe(FullPathZip) = -1) Then Exit Sub
   emailDestino = InputBox("Informe o Email do destinatario", "Envio de Email", rCont!email)
   If Not Vazio(emailDestino) Then
      Call EnviaEmail(emailDestino, FullPathZip)
      DoEvents
   End If
   
   Set rCont = Nothing
   lblAguarde.Visible = False
End Sub

Private Sub Form_Load()
Dim totalRegistros As Long

Set moCombo = New cComboHelper

'DiretorioDestino = appPathApp

sSQL = "SELECT * FROM empresa"
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If totalRegistros >= 1 Then
      If Not r Is Nothing Then
      vFantasia = ValidateNull(r("fantasia"))
      vRazao = r("razao")
      vCnpj = r("cnpj")
      vCaminhoXML = r("DiretorioXML")
      'txtDiretorioXML.Text = ValidateNull(rTabela("caminho"))
      'txtDiretorioXML.Text = App.path & "\ExportarXML"
      txtDiretorioXML.Text = vCaminhoXML & "\ExportarXML"
   End If
Else
    MsgBox "Precisa fazer o cadastro da licença antes dessa operação!", vbInformation, "Aviso do Sistema"
End If

Set moCombo = New cComboHelper
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub



