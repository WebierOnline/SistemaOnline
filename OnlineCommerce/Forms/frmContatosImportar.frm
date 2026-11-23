VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmContatosImportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar contatos do celular"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblArquivo 
      Caption         =   "Arquivo:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
   Begin VB.TextBox txtArquivo 
      Height          =   315
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   150
      Width           =   7020
   End
   Begin ChamaleonBtn.chameleonButton cmdArquivo 
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Escolher..."
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
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblResumo 
      Caption         =   ""
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   9375
   End
   Begin MSFlexGridLib.MSFlexGrid grdContatos 
      Height          =   5415
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   9375
      _ExtentX        =   16537
      _ExtentY        =   9552
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdImportar 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Importar novos"
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
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      Top             =   6720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Fechar"
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
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmContatosImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Importa os contatos do celular (arquivo .csv exportado do Google Contatos / celular) para a tabela Contatos (script 194),
'usada em Promoções por WhatsApp (frmCampanhaZap, "Contatos do celular"). Cada celular válido vira um contato
'(contato com 2 celulares = 2 linhas). Quem já está na tabela não muda (nem o ATIVO). Fixo e número incompleto ficam de fora.

Private Const TITULO As String = "Importar contatos do celular"
Private Const S_NOVO As String = "Novo"
Private Const S_JA As String = "Já cadastrado"

Private mNome() As String
Private mCel() As String
Private mSit() As String
Private mQtd As Long
Private mNovos As Long
Private mIniciou As Boolean

Private Sub Form_Load()
lblResumo.Caption = "Escolha o arquivo .csv com os contatos exportados do celular (Google Contatos > Exportar > CSV do Google)."
ReDim mNome(0): ReDim mCel(0): ReDim mSit(0)
MontarGrid
cmdImportar.Enabled = False
End Sub

Private Sub Form_Activate()
If mIniciou Then Exit Sub
mIniciou = True
If Not Camp_TemContatos() Then
    MsgBox "O banco de dados ainda não tem a tabela de contatos (script 194)." & vbCrLf & vbCrLf & _
           "Atualize o sistema e tente de novo.", vbExclamation, TITULO
    Unload Me
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

Private Sub cmdArquivo_Click()
On Error GoTo Cancelou
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Arquivo de contatos"
CommonDialog1.Filter = "Contatos (*.csv)|*.csv|Todos os arquivos (*.*)|*.*"
CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
txtArquivo.Text = CommonDialog1.FileName
LerArquivo
Exit Sub
Cancelou:
Err.Clear
End Sub

'---------------------------------------------------------------------------
' Leitura do arquivo
'---------------------------------------------------------------------------
Private Sub LerArquivo()
Dim sTexto As String, Linhas As Collection, v As Variant, cab As Variant, sSep As String, i As Long, j As Long, k As Long
Dim nNomes(0 To 5) As Long, nGenerico As Long, nFones() As Long, nQtdFones As Long, sCampo As String
Dim sNome As String, sCel As String, sMotivo As String, vNum As Variant
Dim dicArq As Object, dicBanco As Object, nLinhas As Long, nFixos As Long, nSemFone As Long, nRepetidos As Long, nJa As Long
Dim bAchou As Boolean

ReDim mNome(0): ReDim mCel(0): ReDim mSit(0): mQtd = 0: mNovos = 0
If Len(txtArquivo.Text) = 0 Then Exit Sub
Me.MousePointer = vbHourglass
On Error GoTo Falha
sTexto = LerTextoUtf8(txtArquivo.Text)
If Left$(sTexto, 1) = ChrW$(&HFEFF) Then sTexto = Mid$(sTexto, 2)
i = InStr(sTexto, vbLf)
If i = 0 Then i = Len(sTexto)
If Len(Left$(sTexto, i)) - Len(Replace(Left$(sTexto, i), ";", "")) > Len(Left$(sTexto, i)) - Len(Replace(Left$(sTexto, i), ",", "")) Then sSep = ";" Else sSep = ","
Set Linhas = New Collection
LerCsv sTexto, sSep, Linhas
sTexto = ""
If Linhas.Count < 2 Then Err.Raise vbObjectError + 1, , "O arquivo não tem contatos."

'Cabeçalho: nome (Google: First/Middle/Last Name, File As, Nickname, Organization Name) e telefones (Phone N - Value)
cab = Linhas(1)
For j = 0 To 5: nNomes(j) = -1: Next
nGenerico = -1
ReDim nFones(0)
For j = 0 To UBound(cab)
    sCampo = LCase$(Trim$(cab(j)))
    Select Case sCampo
        Case "first name", "given name": nNomes(0) = j
        Case "middle name", "additional name": nNomes(1) = j
        Case "last name", "family name": nNomes(2) = j
        Case "file as": nNomes(3) = j
        Case "nickname": nNomes(4) = j
        Case "organization name", "organization 1 - name": nNomes(5) = j
        Case "name", "nome", "nome completo", "display name": nGenerico = j
        Case Else
            If sCampo Like "phone*value" Or ((InStr(sCampo, "phone") > 0 Or InStr(sCampo, "telefone") > 0 Or InStr(sCampo, "celular") > 0 Or _
               InStr(sCampo, "mobile") > 0) And InStr(sCampo, "label") = 0 And InStr(sCampo, "type") = 0 And InStr(sCampo, "tipo") = 0) Then
                nQtdFones = nQtdFones + 1
                ReDim Preserve nFones(nQtdFones)
                nFones(nQtdFones) = j
            End If
    End Select
Next
If nQtdFones = 0 Then Err.Raise vbObjectError + 2, , "Não achei a coluna dos telefones no arquivo (ex.: ""Phone 1 - Value"")."

'Celulares que já estão na tabela
Set dicBanco = CreateObject("Scripting.Dictionary")
Set dicArq = CreateObject("Scripting.Dictionary")
Dim rs As ADODB.Recordset
Set rs = dbData.OpenRecordset("SELECT Celular FROM Contatos")
Do While Not rs.EOF
    dicBanco(Trim$("" & rs("Celular"))) = 1
    rs.MoveNext
Loop
rs.Close

For i = 2 To Linhas.Count
    v = Linhas(i)
    If UBound(v) >= 1 Or Len(Trim$(v(0))) > 0 Then
        nLinhas = nLinhas + 1
        sNome = Trim$(Campo(v, nNomes(0)) & " " & Campo(v, nNomes(1)) & " " & Campo(v, nNomes(2)))
        Do While InStr(sNome, "  ") > 0: sNome = Replace(sNome, "  ", " "): Loop
        If Len(sNome) = 0 Then sNome = Campo(v, nGenerico)
        If Len(sNome) = 0 Then sNome = Campo(v, nNomes(3))
        If Len(sNome) = 0 Then sNome = Campo(v, nNomes(4))
        If Len(sNome) = 0 Then sNome = Campo(v, nNomes(5))
        bAchou = False
        For k = 1 To nQtdFones
            For Each vNum In Split(Campo(v, nFones(k)), ":::")
                If Len(Trim$(vNum)) > 0 Then
                    bAchou = True
                    sCel = Camp_NormalizarCelular(CStr(vNum), sMotivo)
                    If Len(sCel) = 0 Then
                        nFixos = nFixos + 1
                    ElseIf dicArq.Exists(sCel) Then
                        nRepetidos = nRepetidos + 1
                    Else
                        dicArq(sCel) = 1
                        mQtd = mQtd + 1
                        ReDim Preserve mNome(mQtd): ReDim Preserve mCel(mQtd): ReDim Preserve mSit(mQtd)
                        If Len(sNome) = 0 Then mNome(mQtd) = sCel Else mNome(mQtd) = Left$(sNome, 160)
                        mCel(mQtd) = sCel
                        If dicBanco.Exists(sCel) Then
                            mSit(mQtd) = S_JA
                            nJa = nJa + 1
                        Else
                            mSit(mQtd) = S_NOVO
                            mNovos = mNovos + 1
                        End If
                    End If
                End If
            Next
        Next
        If Not bAchou Then nSemFone = nSemFone + 1
    End If
Next
Set Linhas = Nothing
OrdenarPorNome
MontarGrid
lblResumo.Caption = nLinhas & " contato(s) no arquivo: " & mNovos & " celular(es) novo(s) para importar, " & nJa & " já cadastrado(s)." & vbCrLf & _
    "Ficam de fora: " & nFixos & " número(s) que não são celular (fixo, incompleto ou DDD errado), " & nRepetidos & " repetido(s), " & _
    nSemFone & " contato(s) sem telefone."
cmdImportar.Enabled = (mNovos > 0)
Me.MousePointer = vbDefault
Exit Sub
Falha:
Me.MousePointer = vbDefault
MsgBox "Não foi possível ler o arquivo." & vbCrLf & vbCrLf & Err.Description, vbExclamation, TITULO
Err.Clear
mQtd = 0: mNovos = 0
MontarGrid
lblResumo.Caption = "Escolha o arquivo .csv com os contatos exportados do celular."
cmdImportar.Enabled = False
End Sub

Private Function Campo(ByRef v As Variant, ByVal j As Long) As String
If j < 0 Then Exit Function
If j > UBound(v) Then Exit Function
Campo = Trim$(Replace(Replace(v(j), vbCr, " "), vbLf, " "))
End Function

'Arquivo texto em UTF-8 (o celular/Google exporta assim)
Private Function LerTextoUtf8(ByVal sArquivo As String) As String
Dim st As Object

Set st = CreateObject("ADODB.Stream")
st.Type = 2            'texto
st.Charset = "utf-8"
st.Open
st.LoadFromFile sArquivo
LerTextoUtf8 = st.ReadText
st.Close
End Function

'CSV com aspas (campo com vírgula, aspas "" e quebra de linha dentro de aspas). Cada linha vai como array de String.
Private Sub LerCsv(ByRef s As String, ByVal sSep As String, ByRef Linhas As Collection)
Dim i As Long, n As Long, ch As String, bAspas As Boolean, sCampo As String, campos() As String, nc As Long

n = Len(s)
i = 1
nc = 0
ReDim campos(0)
Do While i <= n
    ch = Mid$(s, i, 1)
    If bAspas Then
        If ch = """" Then
            If Mid$(s, i + 1, 1) = """" Then
                sCampo = sCampo & """"
                i = i + 1
            Else
                bAspas = False
            End If
        Else
            sCampo = sCampo & ch
        End If
    ElseIf ch = """" Then
        bAspas = True
    ElseIf ch = sSep Then
        ReDim Preserve campos(nc)
        campos(nc) = sCampo
        nc = nc + 1
        sCampo = ""
    ElseIf ch = vbCr Or ch = vbLf Then
        If ch = vbCr And Mid$(s, i + 1, 1) = vbLf Then i = i + 1
        ReDim Preserve campos(nc)
        campos(nc) = sCampo
        Linhas.Add campos
        nc = 0
        sCampo = ""
        ReDim campos(0)
    Else
        sCampo = sCampo & ch
    End If
    i = i + 1
Loop
If nc > 0 Or Len(sCampo) > 0 Then
    ReDim Preserve campos(nc)
    campos(nc) = sCampo
    Linhas.Add campos
End If
End Sub

'Ordem alfabética (inserção: poucos milhares de linhas)
Private Sub OrdenarPorNome()
Dim i As Long, j As Long, a As String, b As String, c As String

For i = 2 To mQtd
    a = mNome(i): b = mCel(i): c = mSit(i)
    j = i - 1
    Do While j >= 1
        If StrComp(mNome(j), a, vbTextCompare) <= 0 Then Exit Do
        mNome(j + 1) = mNome(j): mCel(j + 1) = mCel(j): mSit(j + 1) = mSit(j)
        j = j - 1
    Loop
    mNome(j + 1) = a: mCel(j + 1) = b: mSit(j + 1) = c
Next
End Sub

'---------------------------------------------------------------------------
' Grid
'---------------------------------------------------------------------------
Private Sub MontarGrid()
Dim i As Long, c As Integer

With grdContatos
    .Redraw = False
    .Rows = 1
    .Cols = 3
    .TextMatrix(0, 0) = "Nome"
    .TextMatrix(0, 1) = "Celular"
    .TextMatrix(0, 2) = "Situação"
    For c = 0 To 2
        .ColAlignment(c) = flexAlignLeftCenter
    Next
    .ColWidth(0) = 5100
    .ColWidth(1) = 1800
    .ColWidth(2) = 2100
    .Rows = mQtd + 1
    For i = 1 To mQtd
        .TextMatrix(i, 0) = mNome(i)
        .TextMatrix(i, 1) = mCel(i)
        .TextMatrix(i, 2) = mSit(i)
        If mSit(i) = S_JA Then
            .Row = i
            For c = 0 To 2
                .Col = c
                .CellForeColor = &H808080
            Next
        End If
    Next
    If .Rows > 1 Then .Row = 1: .Col = 0
    .Redraw = True
End With
End Sub

'---------------------------------------------------------------------------
' Gravação
'---------------------------------------------------------------------------
Private Sub cmdImportar_Click()
Dim i As Long, s As String, nLote As Long, nOk As Long, bTrans As Boolean

If mNovos = 0 Then Exit Sub
If MsgBox("Importar " & mNovos & " contato(s) novo(s)?" & vbCrLf & vbCrLf & _
          "Eles entram ativos. Em Promoções por WhatsApp, marque ""Contatos do celular"", clique em Buscar clientes " & _
          "e use Desativar contato para quem não deve receber.", vbQuestion + vbYesNo, TITULO) <> vbYes Then Exit Sub
Me.MousePointer = vbHourglass
On Error GoTo Falha
dbData.Execute "BEGIN TRANSACTION"
bTrans = True
For i = 1 To mQtd
    If mSit(i) = S_NOVO Then
        s = s & IIf(nLote > 0, ", ", "") & "(" & Cob_Q(mNome(i)) & ", " & Cob_Q(mCel(i)) & ")"
        nLote = nLote + 1
        nOk = nOk + 1
    End If
    If nLote > 0 And (nLote = 200 Or i = mQtd) Then
        dbData.Execute "INSERT INTO Contatos (Nome, Celular) VALUES " & s
        s = "": nLote = 0
    End If
Next
dbData.Execute "COMMIT TRANSACTION"
bTrans = False
Me.MousePointer = vbDefault
MsgBox nOk & " contato(s) importado(s).", vbInformation, TITULO
LerArquivo
Exit Sub
Falha:
Me.MousePointer = vbDefault
s = Err.Description
Err.Clear
On Error Resume Next
If bTrans Then dbData.Execute "ROLLBACK TRANSACTION"
Err.Clear
MsgBox "Não foi possível importar os contatos." & vbCrLf & vbCrLf & s, vbExclamation, TITULO
End Sub
