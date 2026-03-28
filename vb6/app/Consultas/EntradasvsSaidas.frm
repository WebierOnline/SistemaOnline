VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form EntradasvsSaidas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COMPARATIVO - ENTRADAS VS SAÍDAS"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14400
   Icon            =   "EntradasvsSaidas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1395
      Left            =   240
      TabIndex        =   21
      Top             =   5340
      Width           =   13875
      Begin VB.ComboBox cboAno 
         Height          =   315
         ItemData        =   "EntradasvsSaidas.frx":23D2
         Left            =   9060
         List            =   "EntradasvsSaidas.frx":23D4
         TabIndex        =   29
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   7260
         TabIndex        =   28
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6360
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cboDescricao 
         Height          =   315
         ItemData        =   "EntradasvsSaidas.frx":23D6
         Left            =   2220
         List            =   "EntradasvsSaidas.frx":23D8
         TabIndex        =   25
         Top             =   600
         Width           =   4695
      End
      Begin VB.ComboBox cboCriterio 
         Height          =   315
         ItemData        =   "EntradasvsSaidas.frx":23DA
         Left            =   180
         List            =   "EntradasvsSaidas.frx":23DC
         TabIndex        =   22
         Top             =   600
         Width           =   1995
      End
      Begin ChamaleonBtn.chameleonButton chameleonButton1 
         Height          =   555
         Left            =   11280
         TabIndex        =   24
         Top             =   360
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "&Exibir"
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
         MICON           =   "EntradasvsSaidas.frx":23DE
         PICN            =   "EntradasvsSaidas.frx":23FA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox mskFim 
         Height          =   315
         Left            =   9060
         TabIndex        =   32
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   315
         Left            =   7500
         TabIndex        =   33
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton cmdCal1 
         Height          =   315
         Left            =   8700
         TabIndex        =   36
         Tag             =   "Calendario"
         Top             =   960
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   8
         TX              =   ""
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "EntradasvsSaidas.frx":418C
         PICN            =   "EntradasvsSaidas.frx":41A8
         PICH            =   "EntradasvsSaidas.frx":64FB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCal2 
         Height          =   315
         Left            =   10260
         TabIndex        =   37
         Tag             =   "Calendario"
         Top             =   960
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   8
         TX              =   ""
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "EntradasvsSaidas.frx":884E
         PICN            =   "EntradasvsSaidas.frx":886A
         PICH            =   "EntradasvsSaidas.frx":ABBD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblFim 
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9045
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblInicio 
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7800
         TabIndex        =   34
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   9120
         TabIndex        =   31
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Left            =   7260
         TabIndex        =   30
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "Critério:"
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
         Left            =   180
         TabIndex        =   23
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   14205
      TabIndex        =   18
      Top             =   60
      Width           =   14235
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Left            =   8640
         TabIndex        =   19
         Top             =   300
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Picture         =   "EntradasvsSaidas.frx":CF10
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "COMPARATIVO - ENTRADAS VS SAÍDAS"
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
         Left            =   1365
         TabIndex        =   20
         Top             =   240
         Width           =   6075
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1275
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Width           =   11595
      Begin VB.ComboBox cboConsAno 
         Height          =   315
         Left            =   8160
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.ComboBox cboConsDescricao 
         Height          =   315
         ItemData        =   "EntradasvsSaidas.frx":128E3
         Left            =   6300
         List            =   "EntradasvsSaidas.frx":128E5
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cboConsulta 
         Height          =   315
         ItemData        =   "EntradasvsSaidas.frx":128E7
         Left            =   120
         List            =   "EntradasvsSaidas.frx":128E9
         TabIndex        =   5
         Top             =   480
         Width           =   1995
      End
      Begin VB.ComboBox cboOrdem 
         Height          =   315
         ItemData        =   "EntradasvsSaidas.frx":128EB
         Left            =   2220
         List            =   "EntradasvsSaidas.frx":128ED
         TabIndex        =   4
         Top             =   480
         Width           =   1995
      End
      Begin VB.ComboBox cboConsRef 
         Height          =   315
         Left            =   8220
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblConsDescricao 
         Caption         =   "Descrição"
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
         Left            =   6300
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Critério:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "Ordem:"
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
         Left            =   2220
         TabIndex        =   9
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblConsRef 
         Caption         =   "Referência:"
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
         Left            =   7140
         TabIndex        =   8
         Top             =   900
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8550
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21061
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "23:00"
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1695
      Left            =   60
      TabIndex        =   1
      Top             =   1080
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   2990
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      Height          =   555
      Left            =   11760
      TabIndex        =   12
      Top             =   7800
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "EntradasvsSaidas.frx":128EF
      PICN            =   "EntradasvsSaidas.frx":1290B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibir 
      Height          =   555
      Left            =   11760
      TabIndex        =   13
      Top             =   7140
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "&Exibir"
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
      MICON           =   "EntradasvsSaidas.frx":1469D
      PICN            =   "EntradasvsSaidas.frx":146B9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid GridEntradaSaida 
      Height          =   1695
      Left            =   120
      TabIndex        =   27
      Top             =   3300
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   2990
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.Label lblValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   12420
      TabIndex        =   17
      Top             =   6660
      Width           =   1605
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      Caption         =   "Valor:"
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
      Left            =   11520
      TabIndex        =   16
      Top             =   6660
      Width           =   795
   End
   Begin VB.Label lblQuant 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9960
      TabIndex        =   15
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      Caption         =   "Quant.:"
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
      Left            =   9180
      TabIndex        =   14
      Top             =   6660
      Width           =   735
   End
End
Attribute VB_Name = "EntradasvsSaidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private moCombo As cComboHelper

Private Sub cboTipo_Change()

End Sub


Private Sub cboCriterio_GotFocus()
cboCriterio.Clear
cboCriterio.AddItem "TODOS"
cboCriterio.AddItem "DESCRIÇÃO"
cboCriterio.AddItem "MENSAL"
cboCriterio.AddItem "PERÍODO"
moCombo.AttachTo cboCriterio
End Sub


Private Sub cboDescricao_GotFocus()
moCombo.AttachTo cboDescricao
   
Dim sSQL As String
Dim r As ADODB.Recordset


    If cboDescricao.ListIndex = -1 Then
        cboDescricao.Clear
        
        sSQL = "SELECT DISTINCT descricao, codigo FROM produtos where ATIVO = 1 ORDER BY descricao;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboDescricao.AddItem ValidateNull(r("descricao"))
            cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
           r.MoveNext
        Loop
    End If


moCombo.AttachTo cboDescricao
End Sub


Private Sub cboDescricao_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboDescricao_LostFocus()
On Error GoTo TrataErro

If cboDescricao.Text = "" Then txtCodigo.Text = "": Exit Sub

If cboDescricao.ListIndex = -1 Then txtCodigo.Text = "": Exit Sub

txtCodigo = cboDescricao.ItemData(cboDescricao.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub chameleonButton1_Click()

'INDICE PARA ORGANIZAR OS DADOS
Dim INDICE As String
Dim sSQL As String
Dim r As ADODB.Recordset
'Dim totalRegistros As Long
'Dim fExibir As Integer

If cboCriterio.Text = "" Then Exit Sub

'Seleciona a ordem dos registros
'If cboOrdem.Text = "NUM DA NOTA" Then
'    INDICE = "notafiscal;"
'ElseIf cboOrdem.Text = "DATA" Then
'    INDICE = "data_entrada;"
'ElseIf cboOrdem.Text = "VALOR" Then
'    INDICE = "valor;"
'ElseIf cboOrdem.Text = "FORNECEDOR" Then
'    INDICE = "fornecedor;"
'Else
'    INDICE = "notafiscal;"
'End If
Dim varEntrada As String
Dim varVenda As String

 
If cboCriterio.Text = "TODOS" Then

ElseIf cboCriterio.Text = "MENSAL" Then
    varEntrada = " (MONTH(produtos_entrada.data_entrada) = " & cboMes.ListIndex + 1 & ") AND (YEAR(.produtos_entrada.data_entrada) = " & cboAno & ")"
    varVenda = " (MONTH(pedidos.data_compra) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos.data_compra) = " & cboAno & ")"
ElseIf cboCriterio.Text = "PERÍODO" Then
    varEntrada = " (Produtos_Entrada.data_entrada >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND ( Produtos_Entrada.data_entrada  <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103))"
    varVenda = " ( pedidos.data_compra >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND ( pedidos.data_compra  <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103))"
End If

    sSQL = "SELECT  PRODUTOS.CODIGO, DESCRICAO, " & _
    "(SELECT SUM(produtos_entrada_itens.QUANT) FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.CODIGO = produtos_entrada_itens.CODIGO_ENTRADA WHERE (produtos_entrada_itens.CODIGO_PRODUTO = Produtos.CODIGO) and " & varEntrada & ") AS varQuantEntrada, " & _
    "(SELECT SUM(pedidos_itens.QUANTIDADE) FROM pedidos INNER JOIN pedidos_itens ON pedidos.COD_PEDIDO = pedidos_itens.COD_PEDIDO WHERE (pedidos_itens.COD_PRODUTO = Produtos.CODIGO) and " & varVenda & "   ) AS varQuantVendida, " & _
    "((SELECT SUM(produtos_entrada_itens.QUANT) FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.CODIGO = produtos_entrada_itens.CODIGO_ENTRADA WHERE (produtos_entrada_itens.CODIGO_PRODUTO = Produtos.CODIGO) and " & varEntrada & ") - (SELECT SUM(pedidos_itens.QUANTIDADE) FROM pedidos INNER JOIN pedidos_itens ON pedidos.COD_PEDIDO = pedidos_itens.COD_PEDIDO WHERE (pedidos_itens.COD_PRODUTO = Produtos.CODIGO) and " & varVenda & " )) as varQuantRestante " & _
    "FROM Produtos " & _
    "ORDER BY PRODUTOS.CODIGO"
    
Set r = dbData.OpenRecordset(sSQL)

'Exibe o resultado
FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing

'lblQuant.Caption = SomaGrid(Grid, 4)
'lblValor.Caption = Format(SomaGrid(Grid, 5), "##,##0.00")
End Sub
Private Sub cboMes_GotFocus()
cboMes.Clear

cboMes.AddItem "Janeiro"
cboMes.AddItem "Fevereiro"
cboMes.AddItem "Março"
cboMes.AddItem "Abril"
cboMes.AddItem "Maio"
cboMes.AddItem "Junho"
cboMes.AddItem "Julho"
cboMes.AddItem "Agosto"
cboMes.AddItem "Setembro"
cboMes.AddItem "Outubro"
cboMes.AddItem "Novembro"
cboMes.AddItem "Dezembro"

moCombo.AttachTo cboMes
End Sub
Private Sub cboano_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboAno.AddItem i
Next

moCombo.AttachTo cboAno
End Sub
Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With GridEntradaSaida
      .Clear
      .Cols = 6
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 900
      .ColWidth(2) = 4000
      .ColWidth(3) = 1100
      .ColWidth(4) = 1100
      .ColWidth(5) = 1100
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "ENTRADAS"
      .TextMatrix(0, 4) = "SAÍDAS"
      .TextMatrix(0, 5) = "ESTOQUE"
      .Redraw = False
      i = 1
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("CODIGO"))
            .TextMatrix(.Rows - 1, 2) = rTabela("DESCRICAO")
            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("varQuantEntrada"))
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("varQuantVendida"))
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("varQuantRestante"))
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 5
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   'lblTotalEntrada.Caption = Format(SomaGrid(Grid, 3), ocMONEY)
   'lblTotalSaida.Caption = Format(SomaGrid(Grid, 4), ocMONEY)
   'lblTotal.Caption = Format(SomaGrid(Grid, 5), ocMONEY)
End Sub


Private Sub chameleonButton2_Click()

End Sub

Private Sub cmdCal1_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty                    'Inicializa a variável

Set fCal = New Calendario      'Cria o form de calendário
fCal.Show vbModal

varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrói a variável

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskInicio = Format(varData, "dd/mm/yyyy")   'Exibe a data no campo
End Sub

Private Sub cmdCal2_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty                    'Inicializa a variável

Set fCal = New Calendario      'Cria o form de calendário
fCal.Show vbModal

varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrói a variável

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskFim = Format(varData, "dd/mm/yyyy")   'Exibe a data no campo
End Sub

Private Sub cmdExibir_Click()
'INDICE PARA ORGANIZAR OS DADOS
Dim INDICE As String
Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long
Dim fExibir As Integer

If cboConsulta.Text = "" Then Exit Sub

'Seleciona a ordem dos registros
If cboOrdem.Text = "NUM DA NOTA" Then
    INDICE = "notafiscal;"
ElseIf cboOrdem.Text = "DATA" Then
    INDICE = "data_entrada;"
ElseIf cboOrdem.Text = "VALOR" Then
    INDICE = "valor;"
ElseIf cboOrdem.Text = "FORNECEDOR" Then
    INDICE = "fornecedor;"
Else
    INDICE = "notafiscal;"
End If

fExibir = 1

If cboConsulta.Text = "TODOS" Then
  sSQL = "SELECT produtos.*, produtos_entrada.*, produtos_entrada_itens.*, produtos_entrada_itens.descricao as varDesc, produtos_entrada.codigo AS varCodEnt, produtos_entrada.data_entrada as varData, produtos_entrada_itens.quant as varQuant " & _
     "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
     "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _
     "ORDER BY varDesc, " & INDICE
  
  'sSQL = "SELECT notafiscal, ref, produtos.fabricante as produtos.FABRICANTE, produtos.tamanho as produtos.TAMANHO, produtos_entrada.codigo AS produtos_entrada.codigo, produtos_entrada.data_entrada AS produtos_entrada.DATA_ENTRADA, " & _
     "produtos_entrada_itens.descricao AS var_desc, produtos_entrada_itens.quant AS produtos_entrada_itens.quant, " & _
     "produtos_entrada_itens.custo AS var_custo, produtos_entrada_itens.frete_compra AS var_frete, " & _
     "produtos_entrada_itens.imposto_valor_compra AS var_impcompra, produtos_entrada_itens.custo_compra AS var_vlrcompra, " & _
     "produtos_entrada_itens.lucro_valor AS var_lucro, produtos_entrada_itens.imposto_valor_venda AS var_impvenda, " & _
     "produtos_entrada_itens.venda AS var_vlrvenda, produtos_entrada.*, produtos_entrada_itens.* " & _
     "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
     "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _
     "ORDER BY produtos_entrada_itens.descricao, " & INDICE

ElseIf cboConsulta.Text = "MENSAL" Then
If Not ExistInList(cboConsDescricao) Then
  ShowMsg "Selecione o mês na lista.", vbExclamation
  Exit Sub
End If

If Not ExistInList(cboConsAno) Then
  ShowMsg "Selecione o ano na lista.", vbExclamation
  Exit Sub
End If

If cboConsAno.Text = "" Or cboConsDescricao.Text = "" Then Exit Sub

  sSQL = "SELECT produtos.*, produtos_entrada.*, produtos_entrada_itens.*, produtos_entrada_itens.descricao as varDesc, produtos_entrada.codigo AS varCodEnt, produtos_entrada.data_entrada as varData, produtos_entrada_itens.quant as varQuant " & _
  "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
  "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _
  "WHERE (MONTH(data_entrada) = " & cboConsDescricao.ListIndex + 1 & ") AND (YEAR(data_entrada) = " & cboConsAno & ") " & _
  "ORDER BY produtos_entrada_itens.descricao, " & INDICE

ElseIf cboConsulta.Text = "DETALHADO" Then

  If cboConsDescricao.Text = "FABRICANTE" Then
  
  sSQL = "SELECT produtos.*, produtos_entrada.*, produtos_entrada_itens.*, produtos_entrada_itens.descricao as varDesc, produtos_entrada.codigo AS varCodEnt, produtos_entrada.data_entrada as varData, produtos_entrada_itens.quant as varQuant " & _
      "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
      "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _
      "WHERE  (fabricante = '" & cboConsAno.Text & "') " & _
      "ORDER BY produtos_entrada_itens.descricao, " & INDICE
  
  ElseIf cboConsDescricao.Text = "REFERENCIA" Then

  sSQL = "SELECT produtos.*, produtos_entrada.*, produtos_entrada_itens.*, produtos_entrada_itens.descricao as varDesc, produtos_entrada.codigo AS varCodEnt, produtos_entrada.data_entrada as varData, produtos_entrada_itens.quant as varQuant " & _
      "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
      "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _
      "WHERE  (REF = '" & cboConsAno.Text & "') " & _
      "ORDER BY produtos_entrada_itens.descricao, " & INDICE
  
  ElseIf cboConsDescricao.Text = "TAMANHO" Then

  sSQL = "SELECT produtos.*, produtos_entrada.*, produtos_entrada_itens.*, produtos_entrada_itens.descricao as varDesc, produtos_entrada.codigo AS varCodEnt, produtos_entrada.data_entrada as varData, produtos_entrada_itens.quant as varQuant " & _
      "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
      "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _
      "WHERE  (TAMANHO = '" & cboConsAno.Text & "') " & _
      "ORDER BY produtos_entrada_itens.descricao, " & INDICE

  ElseIf cboConsDescricao.Text = "LINHA" Then

  sSQL = "SELECT produtos.*, produtos_entrada.*, produtos_entrada_itens.*, produtos_entrada_itens.descricao as varDesc, produtos_entrada.codigo AS varCodEnt, produtos_entrada.data_entrada as varData, produtos_entrada_itens.quant as varQuant " & _
      "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
      "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _
      "WHERE  (CATEGORIA = '" & cboConsAno.Text & "') " & _
      "ORDER BY produtos_entrada_itens.descricao, " & INDICE

  ElseIf cboConsDescricao.Text = "COD. BARRA" Then

  sSQL = "SELECT produtos.*, produtos_entrada.*, produtos_entrada_itens.*, produtos_entrada_itens.descricao as varDesc, produtos_entrada.codigo AS varCodEnt, produtos_entrada.data_entrada as varData, produtos_entrada_itens.quant as varQuant " & _
      "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
      "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _
      "WHERE  (COD_BARRA = '" & cboConsAno.Text & "') " & _
      "ORDER BY produtos_entrada_itens.descricao, " & INDICE

  End If
  
ElseIf cboConsulta.Text = "DETALHADO + MENSAL" Then

   If Not ExistInList(cboConsDescricao) Then
      ShowMsg "Selecione o mês na lista.", vbExclamation
      Exit Sub
   End If
   
   If Not ExistInList(cboConsAno) Then
      ShowMsg "Selecione o ano na lista.", vbExclamation
      Exit Sub
   End If

  If cboConsAno.Text = "" Or cboConsDescricao.Text = "" Or cboConsRef.Text = "" Then Exit Sub

  sSQL = "SELECT produtos.*, produtos_entrada.*, produtos_entrada_itens.*, produtos_entrada_itens.descricao as varDesc, produtos_entrada.codigo AS varCodEnt, produtos_entrada.data_entrada as varData, produtos_entrada_itens.quant as varQuant " & _
      "FROM produtos_entrada INNER JOIN produtos_entrada_itens ON produtos_entrada.codigo = produtos_entrada_itens.codigo_entrada " & _
      "INNER JOIN produtos ON produtos.codigo = produtos_entrada_itens.codigo_produto  " & _
      "WHERE (REF = '" & cboConsRef.Text & "') AND (MONTH(data_entrada) = " & cboConsDescricao.ListIndex + 1 & ") AND (YEAR(data_entrada) = " & cboConsAno & ") " & _
      "ORDER BY produtos_entrada_itens.descricao, " & INDICE
  End If

Set r = dbData.OpenRecordset(sSQL)

'Exibe o resultado
FormatarGrid_entradas r, True

If r.State <> 0 Then r.Close
Set r = Nothing

'lblQuant.Caption = SomaGrid(Grid, 4)
lblValor.Caption = Format(SomaGrid(Grid, 5), "##,##0.00")

End Sub
Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Double
Dim i As Integer, Valor As Currency

Valor = 0
For i = 0 To var_Grid.Rows - 1
   If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
      Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
   End If
Next

SomaGrid = Valor
End Function

Private Sub FormatarGrid_entradas(rTabela As ADODB.Recordset, Optional ByVal Agrupar As Boolean = False)
Dim i As Integer, x As Integer

Dim aux As String, iRow As Long
Dim subtotalQtde As Double
Dim bNovoGrupo As Boolean

  With Grid
     .Clear
     .Cols = 14
     .Rows = 2
     
     .ColWidth(0) = 0
     .ColWidth(1) = 0
     .ColWidth(2) = 900
     .ColWidth(3) = 4250
     .ColWidth(4) = 750
     .ColWidth(5) = 750
     .ColWidth(6) = 750
     .ColWidth(7) = 750
     .ColWidth(8) = 750
     .ColWidth(9) = 750
     .ColWidth(10) = 750
     .ColWidth(11) = 750
     .ColWidth(12) = 750
     .ColWidth(13) = 750
     
     .TextMatrix(0, 1) = "COD"
     .TextMatrix(0, 2) = "DATA"
     .TextMatrix(0, 3) = "PRODUTO"
     .TextMatrix(0, 4) = "QUANT"
     .TextMatrix(0, 5) = "CUSTO"
     .TextMatrix(0, 6) = "FRETE"
     .TextMatrix(0, 7) = "IMP."
     .TextMatrix(0, 8) = "VALOR"
     .TextMatrix(0, 9) = "LUCRO"
     .TextMatrix(0, 10) = "IMP."
     .TextMatrix(0, 11) = "VENDA"
     .TextMatrix(0, 12) = "VENDA"
     .TextMatrix(0, 13) = "VENDA"
     
     'colocar os cabeçalho em negrito
     For x = 0 To .Cols - 1
        .Col = x
        .Row = 0
        .CellFontBold = True
     Next
     
     'ALINHAMENTO
     .ColAlignment(2) = 1
     
     'centralizar o titulo
     For x = 0 To .Cols - 1
        .Row = 0
        .Col = x
        .CellAlignment = flexAlignCenterCenter
     Next
     
     .Redraw = False
     i = 1
     
     'bNovoGrupo = True
     subtotalQtde = 0
     iRow = 1
     
     If Not rTabela Is Nothing Then
        'Atribui o nome do primeiro item do grupo
        'aux = rTabela("var_desc")
        
        Do While Not rTabela.EOF
           'mudar a cor da coluna
           'For i = 1 To .Rows - 1
           '   .Row = i
           '   .Col = 5:
           '   .CellBackColor = &HC0FFFF
           '   .Col = 11:
           '   .CellBackColor = &HC0C0FF
           'Next
           
           If Agrupar Then
              If aux <> rTabela("varDesc") Then
                 .TextMatrix(iRow, 4) = Format$(subtotalQtde, ocPESO)
                 .TextMatrix(.Rows - 1, 3) = rTabela("varDesc")
                 
                 For i = 3 To 4
                    .Row = .Rows - 1
                    .Col = i
                    .CellFontBold = True
                 Next
                 
                 subtotalQtde = 0
                 iRow = .Rows - 1
                 .Rows = .Rows + 1
              End If
           End If
           
           .TextMatrix(.Rows - 1, 1) = rTabela("varCodEnt")
           .TextMatrix(.Rows - 1, 2) = Format$(rTabela("varData"), "dd/mm/yy")
        If tipoEmpresa = 4 Then
           .TextMatrix(.Rows - 1, 3) = "[" & Format$(rTabela("notafiscal"), "000,000") & "] " & rTabela("varDesc") & " /  " & rTabela("produtos.TAMANHO") & " / " & rTabela("produtos.FABRICANTE")
        Else
           .TextMatrix(.Rows - 1, 3) = "[" & Format$(rTabela("notafiscal"), "000,000") & "] " & rTabela("varDesc")
        End If
           .TextMatrix(.Rows - 1, 4) = Format$(rTabela("varQuant"), ocMONEY)
           '.TextMatrix(.Rows - 1, 5) = Format$(rTabela("var_custo"), ocMONEY)
           '.TextMatrix(.Rows - 1, 6) = Format$(rTabela("var_frete"), ocMONEY)
           '.TextMatrix(.Rows - 1, 7) = Format$(rTabela("var_impcompra"), ocMONEY)
           '.TextMatrix(.Rows - 1, 8) = Format$(rTabela("var_vlrcompra"), ocMONEY)
           '.TextMatrix(.Rows - 1, 9) = Format$(rTabela("var_lucro"), ocMONEY)
           '.TextMatrix(.Rows - 1, 10) = Format$(rTabela("var_impvenda"), ocMONEY)
           '.TextMatrix(.Rows - 1, 11) = Format$(rTabela("var_vlrvenda"), ocMONEY)
           'aux = rTabela("var_Desc")
         .TextMatrix(.Rows - 1, 5) = Format$(rTabela("custo"), ocMONEY)
         .TextMatrix(.Rows - 1, 6) = FormatNumber(rTabela("MARGEM_VV"), 2) & "%"
         .TextMatrix(.Rows - 1, 7) = Format$(rTabela("VALOR_VV"), ocMONEY)
         .TextMatrix(.Rows - 1, 8) = FormatNumber(rTabela("MARGEM_VP"), 2) & "%"
         .TextMatrix(.Rows - 1, 9) = Format$(rTabela("VALOR_VP"), ocMONEY)
         .TextMatrix(.Rows - 1, 10) = FormatNumber(rTabela("MARGEM_AV"), 2) & "%"
         .TextMatrix(.Rows - 1, 11) = Format$(rTabela("VALOR_AV"), ocMONEY)
         .TextMatrix(.Rows - 1, 12) = FormatNumber(rTabela("MARGEM_AP"), 2) & "%"
         .TextMatrix(.Rows - 1, 13) = Format$(rTabela("VALOR_AP"), ocMONEY)
           ''bNovoGrupo = False
           subtotalQtde = subtotalQtde + ValidateNull(rTabela("varQuant"))
           
           rTabela.MoveNext
           .Rows = .Rows + 1
           i = i + 1
        Loop
        
        .TextMatrix(iRow, 4) = Format$(subtotalQtde, ocPESO)
        
        For i = 3 To 4
           .Row = .Rows - 1
           .Col = i
           .CellFontBold = True
        Next
     End If
     
     'MUDAR COR DE FONTE DA COLUNA
     For x = 1 To .Rows - 1
        .Row = x
        .Col = 11
        .CellForeColor = &HC0&
        .CellFontBold = True
     Next
     
     .Rows = .Rows - 1
     .Redraw = True
  End With
  
If cboConsulta.Text = "DETALHADO + MENSAL" Then
    lblQuant.Caption = Format$(subtotalQtde, ocPESO)
Else
    lblQuant.Caption = Format$(0, ocPESO)
End If

  lblValor.Caption = Format(SomaGrid(Grid, 11), "##,##0.000")
End Sub


Private Sub cboConsulta_GotFocus()
cboConsulta.Clear
'cboConsAno.Clear
cboConsulta.AddItem "TODOS"
cboConsulta.AddItem "MENSAL"
cboConsulta.AddItem "DETALHADO"
cboConsulta.AddItem "DETALHADO + MENSAL"
moCombo.AttachTo cboConsulta
End Sub



Private Sub cboConsDescricao_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vMes As Integer

If cboConsulta.Text = "PRODUTO" Then
   cboConsDescricao.Clear
   
   sSQL = "SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboConsDescricao.AddItem ValidateNull(r("descricao"))
      cboConsDescricao.ItemData(cboConsDescricao.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   moCombo.AttachTo cboConsDescricao
ElseIf cboConsulta.Text = "FORNECEDOR" Then
   cboConsDescricao.Clear
   
   sSQL = "SELECT DISTINCT razao FROM fornecedor ORDER BY razao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboConsDescricao.AddItem r("razao")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   
   moCombo.AttachTo cboConsDescricao
ElseIf cboConsulta.Text = "MENSAL" Then

   
   cboConsDescricao.Clear
   
   For vMes = 1 To 12
      cboConsDescricao.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboConsDescricao
ElseIf cboConsulta.Text = "DETALHADO" Then
    cboConsDescricao.Clear
    cboConsDescricao.AddItem "FABRICANTE"
    cboConsDescricao.AddItem "REFERENCIA"
    cboConsDescricao.AddItem "LINHA"
    cboConsDescricao.AddItem "TAMANHO"
    cboConsDescricao.AddItem "COD. BARRA"

ElseIf cboConsulta.Text = "DETALHADO + MENSAL" Then
   cboConsDescricao.Clear
   
   For vMes = 1 To 12
      cboConsDescricao.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboConsDescricao
End If
End Sub
Private Sub cboConsAno_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
        


If cboConsulta.Text = "DETALHADO" Then
    cboConsAno.Clear
    
    If cboConsDescricao.Text = "FABRICANTE" Then
        sSQL = "SELECT DISTINCT fabricante FROM produtos ORDER BY fabricante;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsAno.AddItem ValidateNull(r("fabricante"))
           r.MoveNext
        Loop
    ElseIf cboConsDescricao.Text = "REFERENCIA" Then
        sSQL = "SELECT DISTINCT REF FROM produtos ORDER BY REF;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsAno.AddItem ValidateNull(r("REF"))
           r.MoveNext
        Loop
    ElseIf cboConsDescricao.Text = "LINHA" Then
        sSQL = "SELECT DISTINCT CATEGORIA FROM produtos ORDER BY CATEGORIA;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsAno.AddItem ValidateNull(r("CATEGORIA"))
           r.MoveNext
        Loop
    ElseIf cboConsDescricao.Text = "TAMANHO" Then
        sSQL = "SELECT DISTINCT TAMANHO FROM produtos ORDER BY TAMANHO;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsAno.AddItem ValidateNull(r("TAMANHO"))
           r.MoveNext
        Loop
    ElseIf cboConsDescricao.Text = "COD. BARRA" Then
        sSQL = "SELECT DISTINCT COD_BARRA FROM produtos ORDER BY COD_BARRA;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboConsAno.AddItem ValidateNull(r("COD_BARRA"))
           r.MoveNext
        Loop
    End If
End If

    If cboConsulta.Text = "MENSAL" Or cboConsulta.Text = "DETALHADO + MENSAL" Then
        Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
        Dim i As Integer
        
        'Calcula o intervalo de anos
        iAno = Year(Date)
        FirstYear = iAno - 2
        LastYear = iAno + 2
        
        cboConsAno.Clear
        
        For i = FirstYear To LastYear
           cboConsAno.AddItem i
        Next
    End If


moCombo.AttachTo cboConsAno
End Sub
Private Sub cboConsRef_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
        
    cboConsRef.Clear

    sSQL = "SELECT DISTINCT REF FROM produtos ORDER BY REF;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       cboConsRef.AddItem ValidateNull(r("REF"))
       r.MoveNext
    Loop

moCombo.AttachTo cboConsRef
End Sub

Private Sub cboOrdem_GotFocus()
cboOrdem.Clear
cboOrdem.AddItem "DATA"
cboOrdem.AddItem "NUM DA NOTA"
cboOrdem.AddItem "VALOR"
cboOrdem.AddItem "FORNECEDOR"
moCombo.AttachTo cboOrdem
End Sub

Private Sub cmdImprimir_Click()
Set REL_Prod_Entrada_Produto.Relatorio.Recordset = r
REL_Prod_Entrada_Produto.dfQuant.Caption = lblQuant.Caption
REL_Prod_Entrada_Produto.dfBruto.Caption = lblValor.Caption

If cboConsulta.Text = "MENSAL" Then
   REL_Prod_Entrada_Produto.dfTipo.Caption = "Tipo: Mês = " & cboConsDescricao.Text & "/" & cboConsAno.Text
ElseIf cboConsulta.Text = "PRODUTO" Then
   REL_Prod_Entrada_Produto.dfTipo.Caption = "Tipo: Produto = " & cboConsDescricao.Text & ""
ElseIf cboConsulta.Text = "FORNECEDOR" Then
   REL_Prod_Entrada_Produto.dfTipo.Caption = "Tipo: Fornecedor = " & cboConsDescricao.Text & ""
ElseIf cboConsulta.Text = "NOTA FISCAL" Then
   REL_Prod_Entrada_Produto.dfTipo.Caption = "Tipo: Nota Fiscal Nº " & cboConsDescricao.Text & ""
Else
   REL_Prod_Entrada_Produto.dfTipo.Caption = "Tipo: Todas as notas"
End If

REL_Prod_Entrada_Produto.Relatorio.NomeImpressora = var_Impressora
REL_Prod_Entrada_Produto.Relatorio.Ativar
Unload REL_Prod_Entrada_Produto
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub


Private Sub mskFim_KeyPress(KeyAscii As Integer)
mskFim.Mask = "##/##/##"
End Sub


Private Sub mskInicio_KeyPress(KeyAscii As Integer)
mskInicio.Mask = "##/##/##"
End Sub


