VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportarIBPT 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Tabela IBPT"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8850
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8850
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgArquivo 
      Left            =   8220
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraMain 
      Caption         =   "Arquivo CSV"
      Height          =   1095
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   8760
      Begin VB.TextBox txtArquivo 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   7560
      End
      Begin VB.CommandButton cmdLocalizar 
         Caption         =   "..."
         Height          =   315
         Left            =   7800
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame fraVersao 
      Caption         =   "Versão"
      Height          =   975
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   8760
      Begin VB.Label lblVersaoDB 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblVersaoCSV 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6240
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Versão no banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Versão no CSV:"
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "Importar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   6900
      TabIndex        =   3
      Top             =   2160
      Width           =   1920
   End
   Begin VB.Label lblProgresso 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00007700&
      Height          =   375
      Left            =   60
      TabIndex        =   10
      Top             =   2700
      Visible         =   0   'False
      Width           =   8760
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   60
      TabIndex        =   11
      Top             =   3120
      Width           =   8760
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vigência Fim da tabela:"
      Height          =   195
      Left            =   2520
      TabIndex        =   12
      Top             =   2280
      Width           =   1650
   End
   Begin VB.Label lblVigFim 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "frmImportarIBPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sArquivoCSV As String
Dim sVersaoCSV  As String
Dim sVigFimCSV  As String

Private Sub Form_Load()
    CarregarVersaoDB
End Sub

Private Sub CarregarVersaoDB()
    Dim sVer As String
    Dim sVig As String
    sVer = SQLExecutaRetorno("SELECT TOP 1 versao    FROM TabelaIBPT", "versao", "")
    sVig = SQLExecutaRetorno("SELECT TOP 1 vigenciafim FROM TabelaIBPT", "vigenciafim", "")
    If sVer = "" Then
        lblVersaoDB.Caption = "(vazia)"
        lblVigFim.Caption = ""
    Else
        lblVersaoDB.Caption = sVer
        If sVig <> "" Then
            lblVigFim.Caption = Format(CDate(sVig), "dd/mm/yyyy")
            If CDate(sVig) <= Date + 15 Then lblVigFim.ForeColor = vbRed
        End If
    End If
End Sub

Private Sub cmdLocalizar_Click()
    On Error GoTo CancelaBrowse
    dlgArquivo.Filter = "Arquivos CSV (*.csv)|*.csv"
    dlgArquivo.FilterIndex = 1
    dlgArquivo.DialogTitle = "Selecionar Tabela IBPT"
    dlgArquivo.Flags = &H1000        'OFN_FILEMUSTEXIST
    dlgArquivo.ShowOpen
    sArquivoCSV = dlgArquivo.FileName
    If sArquivoCSV = "" Then Exit Sub
    txtArquivo.Text = sArquivoCSV
    LerVersaoCSV
    Exit Sub
CancelaBrowse:
End Sub

Private Sub LerVersaoCSV()
    Dim iFile As Integer
    Dim sLinha As String
    Dim aCampos() As String
    iFile = FreeFile
    Open sArquivoCSV For Input As #iFile
    Line Input #iFile, sLinha  ' cabecalho - ignora
    If Not EOF(iFile) Then
        Line Input #iFile, sLinha
        aCampos = Split(sLinha, ";")
        If UBound(aCampos) >= 11 Then
            sVersaoCSV = Trim(aCampos(11))
            sVigFimCSV = Trim(aCampos(9))
        End If
    End If
    Close #iFile
    lblVersaoCSV.Caption = sVersaoCSV
    cmdImportar.Enabled = (sVersaoCSV <> "")
End Sub

Private Function VersaoMaisNova(vCSV As String, vDB As String) As Boolean
    If vDB = "" Or vDB = "(vazia)" Then VersaoMaisNova = True: Exit Function
    If vCSV = vDB Then VersaoMaisNova = False: Exit Function
    Dim aCSV() As String, aDB() As String
    aCSV = Split(vCSV, ".")
    aDB = Split(vDB, ".")
    If UBound(aCSV) < 2 Or UBound(aDB) < 2 Then
        VersaoMaisNova = (vCSV > vDB): Exit Function
    End If
    Dim nAnoCSV As Integer, nAnoDB As Integer
    Dim nSeqCSV As Integer, nSeqDB As Integer
    nAnoCSV = Val(aCSV(0)): nAnoDB = Val(aDB(0))
    nSeqCSV = Val(aCSV(1)): nSeqDB = Val(aDB(1))
    If nAnoCSV <> nAnoDB Then VersaoMaisNova = (nAnoCSV > nAnoDB): Exit Function
    If nSeqCSV <> nSeqDB Then VersaoMaisNova = (nSeqCSV > nSeqDB): Exit Function
    VersaoMaisNova = (UCase(aCSV(2)) > UCase(aDB(2)))
End Function

Private Sub cmdImportar_Click()
    Dim sVersaoDB As String
    sVersaoDB = lblVersaoDB.Caption

    If Not VersaoMaisNova(sVersaoCSV, sVersaoDB) Then
        lblStatus.Caption = "A tabela no banco (v" & sVersaoDB & ") já está atualizada. " & _
                            "A versão do CSV (" & sVersaoCSV & ") não é mais nova."
        Exit Sub
    End If

    Dim resp As Integer
    resp = MsgBox("Serão apagados todos os registros atuais e importados " & _
                  "os dados da versão " & sVersaoCSV & "." & vbCr & _
                  "Deseja continuar?", vbQuestion + vbYesNo, "Confirmar Importação")
    If resp <> vbYes Then Exit Sub

    cmdImportar.Enabled = False
    cmdLocalizar.Enabled = False
    cmdFechar.Enabled = False
    lblProgresso.Visible = True
    lblStatus.Caption = ""

    On Error GoTo ErrImport

    ' Limpa tabela
    lblProgresso.Caption = "Limpando tabela..."
    DoEvents
    dbData.Execute "DELETE FROM TabelaIBPT"

    ' Le CSV e insere
    Dim iFile   As Integer
    Dim sLinha  As String
    Dim aCampos() As String
    Dim nTotal  As Long
    Dim nLinha  As Long
    Dim sSQL    As String

    ' Conta linhas para progresso
    iFile = FreeFile
    Open sArquivoCSV For Input As #iFile
    nTotal = 0
    Do While Not EOF(iFile)
        Line Input #iFile, sLinha
        nTotal = nTotal + 1
    Loop
    Close #iFile
    nTotal = nTotal - 1  ' desconta cabecalho

    iFile = FreeFile
    Open sArquivoCSV For Input As #iFile
    Line Input #iFile, sLinha  ' pula cabecalho
    nLinha = 0

    dbData.Execute "BEGIN TRANSACTION"

    Do While Not EOF(iFile)
        Line Input #iFile, sLinha
        sLinha = Trim(sLinha)
        If sLinha = "" Then GoTo ProxLinha
        aCampos = Split(sLinha, ";")
        If UBound(aCampos) < 12 Then GoTo ProxLinha

        Dim sCodigo   As String, sEx        As String
        Dim sTipo     As String, sDesc      As String
        Dim sNacFed   As String, sImpFed    As String
        Dim sEstadual As String, sMunicipal As String
        Dim sVigIni   As String, sVigFim    As String
        Dim sChave    As String, sVersao    As String
        Dim sFonte    As String

        sCodigo = Trim(aCampos(0))
        sEx = Trim(aCampos(1)): If sEx = "" Then sEx = "0"
        sTipo = Trim(aCampos(2)): If sTipo = "" Then sTipo = "0"
        sDesc = Trim(aCampos(3))
        ' Remove aspas da descricao
        If Left(sDesc, 1) = Chr(34) Then sDesc = Mid(sDesc, 2)
        If Right(sDesc, 1) = Chr(34) Then sDesc = Left(sDesc, Len(sDesc) - 1)
        sDesc = Replace(sDesc, "'", "''")
        sNacFed = Trim(aCampos(4)):    If sNacFed = "" Then sNacFed = "0"
        sImpFed = Trim(aCampos(5)):    If sImpFed = "" Then sImpFed = "0"
        sEstadual = Trim(aCampos(6)):  If sEstadual = "" Then sEstadual = "0"
        sMunicipal = Trim(aCampos(7)): If sMunicipal = "" Then sMunicipal = "0"
        sVigIni = Trim(aCampos(8))
        sVigFim = Trim(aCampos(9))
        sChave = Trim(aCampos(10))
        sVersao = Trim(aCampos(11))
        sFonte = Trim(aCampos(12))

        ' Converte datas dd/mm/yyyy -> SQL CONVERT(date,...,103)
        Dim sIniSQL As String, sFimSQL As String
        sIniSQL = IIf(sVigIni <> "", "CONVERT(date,'" & sVigIni & "',103)", "NULL")
        sFimSQL = IIf(sVigFim <> "", "CONVERT(date,'" & sVigFim & "',103)", "NULL")

        sSQL = "INSERT INTO TabelaIBPT " & _
               "(codigo,ex,tipo,descricao,nacionalfederal,importadosfederal," & _
               "estadual,municipal,vigenciainicio,vigenciafim,chave,versao,fonte) VALUES (" & _
               "'" & sCodigo & "'," & _
               "'" & sEx & "'," & _
               sTipo & "," & _
               "'" & sDesc & "'," & _
               sNacFed & "," & sImpFed & "," & sEstadual & "," & sMunicipal & "," & _
               sIniSQL & "," & sFimSQL & "," & _
               "'" & sChave & "'," & _
               "'" & sVersao & "'," & _
               "'" & Replace(sFonte, "'", "''") & "')"

        dbData.Execute sSQL
        nLinha = nLinha + 1

        If nLinha Mod 100 = 0 Then
            lblProgresso.Caption = "Importando: " & nLinha & " de " & nTotal & " registros..."
            DoEvents
        End If
ProxLinha:
    Loop

    Close #iFile
    dbData.Execute "COMMIT TRANSACTION"

    ' Sincroniza tbNCM: atualiza aliquotas dos NCMs ja existentes
    lblProgresso.Caption = "Sincronizando tbNCM..."
    DoEvents
    dbData.Execute "UPDATE N SET " & _
                   "    N.descricao         = I.descricao, " & _
                   "    N.nacionalfederal   = I.nacionalfederal, " & _
                   "    N.importadosfederal = I.importadosfederal, " & _
                   "    N.estadual          = I.estadual, " & _
                   "    N.municipal         = I.municipal " & _
                   "FROM tbNCM N " & _
                   "INNER JOIN TabelaIBPT I ON I.codigo = N.NCM AND I.ex = '0'"

    ' Insere NCMs novos que ainda nao existem em tbNCM
    dbData.Execute "INSERT INTO tbNCM (NCM, descricao, nacionalfederal, importadosfederal, estadual, municipal) " & _
                   "SELECT I.codigo, I.descricao, I.nacionalfederal, I.importadosfederal, I.estadual, I.municipal " & _
                   "FROM TabelaIBPT I " & _
                   "WHERE I.ex = '0' " & _
                   "AND NOT EXISTS (SELECT 1 FROM tbNCM N WHERE N.NCM = I.codigo)"

    lblProgresso.Caption = "Concluído! " & nLinha & " registros importados."
    lblStatus.Caption = "Tabela IBPT versão " & sVersaoCSV & " importada com sucesso!" & vbCr & _
                        "Registros: " & nLinha & "  |  Vigência: " & sVigIni & " a " & sVigFim
    CarregarVersaoDB
    cmdFechar.Enabled = True
    cmdLocalizar.Enabled = True
    Exit Sub

ErrImport:
    Dim sErr As String
    sErr = Err.Description
    On Error Resume Next
    Close #iFile
    dbData.Execute "ROLLBACK TRANSACTION"
    lblProgresso.Caption = "Erro na importação!"
    lblProgresso.BackColor = vbRed
    lblStatus.Caption = "ERRO: " & sErr
    cmdImportar.Enabled = True
    cmdLocalizar.Enabled = True
    cmdFechar.Enabled = True
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub
