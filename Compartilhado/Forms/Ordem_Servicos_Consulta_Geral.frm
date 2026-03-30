VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Ordem_Servicos_Consulta_Geral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORDEM DE SERVIÇO"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   Icon            =   "Ordem_Servicos_Consulta_Geral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   9825
      TabIndex        =   43
      Top             =   60
      Width           =   9855
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "Ordem_Servicos_Consulta_Geral.frx":23D2
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ORDEM DE SERVIÇOS"
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
         Left            =   1140
         TabIndex        =   44
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   60
      ScaleHeight     =   2235
      ScaleWidth      =   9795
      TabIndex        =   9
      ToolTipText     =   "Imprimir"
      Top             =   1020
      Width           =   9855
      Begin VB.Frame Frame1 
         Caption         =   "Indice"
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
         Height          =   1455
         Left            =   1920
         TabIndex        =   2
         Top             =   60
         Width           =   1245
         Begin VB.OptionButton optINDStatus 
            Caption         =   "Status"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1200
            Width           =   915
         End
         Begin VB.OptionButton optINDValor 
            Caption         =   "Valor"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   960
            Width           =   915
         End
         Begin VB.OptionButton optINDData 
            Caption         =   "Data"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   240
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optINDCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   480
            Width           =   915
         End
         Begin VB.OptionButton optINDForma 
            Caption         =   "Forma"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   720
            Width           =   915
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Critérios"
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
         Height          =   1455
         Left            =   4560
         TabIndex        =   20
         Top             =   60
         Width           =   5175
         Begin VB.Frame Frame2 
            Height          =   900
            Left            =   300
            TabIndex        =   21
            Top             =   300
            Width           =   4695
            Begin VB.ComboBox cboMes 
               Height          =   315
               Left            =   120
               TabIndex        =   34
               Top             =   390
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.ComboBox cboAno 
               Height          =   315
               Left            =   1500
               Sorted          =   -1  'True
               TabIndex        =   33
               Top             =   390
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.TextBox txtCodigo 
               Height          =   315
               Left            =   120
               TabIndex        =   31
               Top             =   390
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboMecanico 
               Height          =   315
               Left            =   120
               TabIndex        =   24
               Top             =   390
               Visible         =   0   'False
               Width           =   3525
            End
            Begin VB.ComboBox cboCliente 
               Height          =   315
               Left            =   120
               TabIndex        =   22
               Top             =   390
               Visible         =   0   'False
               Width           =   4425
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   315
               Left            =   120
               TabIndex        =   26
               Top             =   390
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskFim 
               Height          =   315
               Left            =   1500
               TabIndex        =   27
               Top             =   390
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin VB.Label lblMes 
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   180
               Left            =   120
               TabIndex        =   36
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblAno 
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   180
               Left            =   1500
               TabIndex        =   35
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblCodigo 
               BackStyle       =   0  'Transparent
               Caption         =   "Código:"
               Height          =   180
               Left            =   120
               TabIndex        =   32
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblInicio 
               BackStyle       =   0  'Transparent
               Caption         =   "Data inicial:"
               Height          =   180
               Left            =   120
               TabIndex        =   30
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblFim 
               BackStyle       =   0  'Transparent
               Caption         =   "Data final:"
               Height          =   180
               Left            =   1500
               TabIndex        =   29
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblAte 
               BackStyle       =   0  'Transparent
               Caption         =   "até"
               Height          =   240
               Left            =   1200
               TabIndex        =   28
               Top             =   480
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label lblMecanico 
               BackStyle       =   0  'Transparent
               Caption         =   "Mecânico:"
               Height          =   180
               Left            =   120
               TabIndex        =   25
               Top             =   180
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label lblClientes 
               BackStyle       =   0  'Transparent
               Caption         =   "Clientes:"
               Height          =   180
               Left            =   120
               TabIndex        =   23
               Top             =   180
               Visible         =   0   'False
               Width           =   915
            End
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Exibir"
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
         Height          =   1455
         Left            =   3240
         TabIndex        =   3
         Top             =   60
         Width           =   1245
         Begin VB.OptionButton optImprimirAprazo 
            Caption         =   "À &Prazo"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   720
            Width           =   915
         End
         Begin VB.OptionButton optImprimirAvista 
            Caption         =   "À &Vista"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   480
            Width           =   915
         End
         Begin VB.OptionButton optImprimirTodas 
            Caption         =   "&Todas"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Pesquisar :"
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
         Height          =   1455
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   1785
         Begin VB.OptionButton optClientes 
            Caption         =   "&Clientes"
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
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   465
            Width           =   1275
         End
         Begin VB.OptionButton optMecanico 
            Caption         =   "&Mecânico"
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
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optDatas 
            Caption         =   "&Entre Datas"
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
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   690
            Width           =   1395
         End
         Begin VB.OptionButton optCodigo 
            Caption         =   "Código"
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
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   915
            Width           =   1275
         End
         Begin VB.OptionButton optMes 
            Caption         =   "Por &Mês"
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
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1140
            Width           =   1275
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdLocalizar 
         Height          =   615
         Left            =   5880
         TabIndex        =   13
         Top             =   1560
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1085
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
         MICON           =   "Ordem_Servicos_Consulta_Geral.frx":85DE
         PICN            =   "Ordem_Servicos_Consulta_Geral.frx":85FA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSair 
         Height          =   615
         Left            =   8520
         TabIndex        =   14
         Top             =   1560
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1085
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
         MICON           =   "Ordem_Servicos_Consulta_Geral.frx":8ED4
         PICN            =   "Ordem_Servicos_Consulta_Geral.frx":8EF0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImprimir 
         Height          =   615
         Left            =   7200
         TabIndex        =   18
         Top             =   1560
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Imprimir"
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
         MICON           =   "Ordem_Servicos_Consulta_Geral.frx":920A
         PICN            =   "Ordem_Servicos_Consulta_Geral.frx":9226
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5085
      Left            =   60
      TabIndex        =   0
      Top             =   3360
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8969
      _Version        =   393216
      Cols            =   7
      BackColorBkg    =   -2147483633
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   45
      Top             =   8880
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13309
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "23:24"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Left            =   7740
      TabIndex        =   42
      Top             =   8520
      Width           =   510
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4020
      TabIndex        =   19
      Top             =   8520
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   8280
      TabIndex        =   12
      Top             =   8520
      Width           =   1635
   End
   Begin VB.Label lblQtda 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   1140
      TabIndex        =   11
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade:"
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
      TabIndex        =   10
      Top             =   8520
      Width           =   1050
   End
End
Attribute VB_Name = "Ordem_Servicos_Consulta_Geral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper
Private printSQL As String

Dim posX As Single

Private Sub Limpar_Objetos()
   cboMecanico.Text = ""
   mskInicio.Mask = ""
   mskInicio.Text = ""
   mskFim.Mask = ""
   mskFim.Text = ""
   txtCodigo.Text = ""
   cboCliente.Text = ""
   cboMES.Text = ""
   cboAno.Text = ""
End Sub

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
   
   'For i = iAno To FirstYear Step -1
   '   cboAno.AddItem i
   'Next
   
   'iAno = iAno + 1
   'For i = ANO To LastYear
   '   cboAno.AddItem i
   'Next
   
   moCombo.AttachTo cboAno
End Sub

Private Sub cboAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub cboCliente_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboCliente.Clear
   
   sSQL = "SELECT nome, codigo FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCliente.AddItem r("nome")
      cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboCliente
End Sub

Private Sub cboCliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub cboMes_GotFocus()
   cboMES.Clear
   cboMES.AddItem "Janeiro"
   cboMES.AddItem "Fevereiro"
   cboMES.AddItem "Março"
   cboMES.AddItem "Abril"
   cboMES.AddItem "Maio"
   cboMES.AddItem "Junho"
   cboMES.AddItem "Julho"
   cboMES.AddItem "Agosto"
   cboMES.AddItem "Setembro"
   cboMES.AddItem "Outubro"
   cboMES.AddItem "Novembro"
   cboMES.AddItem "Dezembro"
   moCombo.AttachTo cboMES
End Sub

Private Sub cboMes_LostFocus()
   cboAno.SetFocus
End Sub

Private Sub cboMecanico_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboMecanico.Clear
   
   sSQL = "SELECT nome, codigo FROM funcionario WHERE (cargo IN ('MECANICO', 'AUX. MECANICO')) ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboMecanico.AddItem r("nome")
      cboMecanico.ItemData(cboMecanico.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboMecanico
End Sub

Private Sub cboMecanico_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub cmdImprimir_Click()
   'colocar o nome da maquina na barra de status
   Dim r As ADODB.Recordset
   Dim var_Impressora As String
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   Set oIni = Nothing
   
   If optMecanico.Value = False And optClientes.Value = False And optDatas.Value = False And optCodigo.Value = False And optMes.Value = False Then Exit Sub
   cmdLocalizar_Click
   
   Me.Hide
   
   Set r = dbData.OpenRecordset(printSQL)
   
   Set REL_OS_Consulta_Geral.Relatorio.Recordset = r
   REL_OS_Consulta_Geral.dfQuant.Caption = "QUANTIDADE: " & lblQtda.Caption
   REL_OS_Consulta_Geral.dfTotal.Caption = "TOTAL: " & lblTotal.Caption
   
   If optImprimirTodas.Value = True Then
      REL_OS_Consulta_Geral.lblTitulo.Caption = "RELATÓRIO DE OS - TODAS"
   ElseIf optImprimirAvista.Value = True Then
      REL_OS_Consulta_Geral.lblTitulo.Caption = "RELATÓRIO DE OS - À VISTA"
   ElseIf optImprimirAprazo.Value = True Then
      REL_OS_Consulta_Geral.lblTitulo.Caption = "RELATÓRIO DE OS - À PRAZO"
   End If
   
   If optMecanico.Value = True Then
      REL_OS_Consulta_Geral.dfTipo.Caption = "Tipo: Mecanico = " & cboMecanico.Text & ""
   ElseIf optClientes.Value = True Then
      REL_OS_Consulta_Geral.dfTipo.Caption = "Tipo: Cliente"
   ElseIf optDatas.Value = True Then
      REL_OS_Consulta_Geral.dfTipo.Caption = "Tipo: Intervalo de " & mskInicio.Text & " à " & mskFim.Text
   ElseIf optCodigo.Value = True Then
      REL_OS_Consulta_Geral.dfTipo.Caption = "Tipo: Data = " & txtCodigo.Text & ""
   ElseIf optMes.Value = True Then
      REL_OS_Consulta_Geral.dfTipo.Caption = "Tipo: Mês = " & cboMES.Text & "/" & cboAno.Text
   End If
   
   REL_OS_Consulta_Geral.Relatorio.NomeImpressora = var_Impressora
   REL_OS_Consulta_Geral.Relatorio.Ativar
   Unload REL_OS_Consulta_Geral
   
   Me.Show 1
End Sub

Public Sub cmdLocalizar_Click()
   If optMecanico.Value = False And optClientes.Value = False And optDatas.Value = False And optCodigo.Value = False And optMes.Value = False Then Exit Sub
   If optMecanico.Value = True And cboMecanico.Text = "" Then Exit Sub
   If optClientes.Value = True And cboCliente.Text = "" Then Exit Sub
   If optCodigo.Value = True And txtCodigo.Text = "" Then Exit Sub
   If optMes.Value = True And cboMES.Text = "" Or cboAno.Text = "" Then Exit Sub
   
   If optImprimirTodas.Value = False And optImprimirAvista.Value = False And optImprimirAprazo.Value = False Then optImprimirTodas.Value = True
   
   Dim INDICE As String       'INDICE PARA ORGANIZAR OS DADOS
   Dim Tipo As String         'FORMA DE PAGAMENTO
   
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   If optINDData.Value = True Then
      INDICE = "data_entrada"
   ElseIf optINDCliente.Value = True Then
      INDICE = "nome"
   ElseIf optINDForma.Value = True Then
      INDICE = "tipo_pagamento"
   ElseIf optINDValor.Value = True Then
      INDICE = "total"
   ElseIf optINDStatus.Value = True Then
      INDICE = "os.status"
   End If
   
   If optImprimirTodas.Value = True Then
      Tipo = ""
   ElseIf optImprimirAvista.Value = True Then
      Tipo = " AND (tipo_pagamento = 'À Vista') "
   ElseIf optImprimirAprazo.Value = True Then
      Tipo = " AND (tipo_pagamento = 'À Prazo') "
   End If
   
   If optMecanico.Value = True Then
      sSQL = "SELECT cliente.*, os.*, os.status AS var_status FROM cliente INNER JOIN os ON cliente.codigo = os.cod_cliente " & _
         "WHERE (cod_mecanico = " & cboMecanico.ItemData(cboMecanico.ListIndex) & ") " & Tipo & " ORDER BY " & INDICE
            
      'SOMAR TODOS =======================
      'Data2.RecordSource = "SELECT SUM(TOTAL)AS VALOR_TOTAL FROM OS WHERE (COD_MECANICO = " & cboMecanico.ItemData(cboMecanico.ListIndex) & ") " & TIPO & ""
      'Data2.Refresh
      'lblTotal.Caption = FormatCurrency(RS!VALOR_TOTAL)
      
   ElseIf optClientes.Value = True Then
      sSQL = "SELECT cliente.*, os.*, os.status AS var_status FROM cliente INNER JOIN os ON cliente.codigo = os.cod_cliente " & _
         "WHERE (cod_cliente = " & cboCliente.ItemData(cboCliente.ListIndex) & ") " & Tipo & " ORDER BY " & INDICE
      
      'SOMAR TODOS =======================
      'Data2.RecordSource = "SELECT SUM(TOTAL)AS VALOR_TOTAL FROM OS WHERE (COD_CLIENTE = " & cboCliente.ItemData(cboCliente.ListIndex) & ") " & TIPO & ""
      'Data2.Refresh
      'lblTotal.Caption = FormatCurrency(RS!VALOR_TOTAL)
      
   ElseIf optDatas.Value = True Then
      If Not IsDate(mskInicio) Or Not IsDate(mskFim) Then Exit Sub
      sSQL = "SELECT cliente.*, os.*, os.status AS var_status FROM cliente INNER JOIN os ON cliente.codigo = os.cod_cliente " & _
         "WHERE (os.data_entrada >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (os.data_entrada <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA_EUA) & "', 103)) " & Tipo & " ORDER BY " & INDICE
      
      'SOMAR TODOS =======================
      'Data2.RecordSource = "SELECT SUM(TOTAL)AS VALOR_TOTAL FROM OS WHERE OS.DATA_ENTRADA BETWEEN #" & Format(mskInicio.Text, "MM/DD/YY") & "# and #" & Format(mskFim.Text, "MM/DD/YY") & "# " & TIPO & ""
      'Data2.Refresh
      'lblTotal.Caption = FormatCurrency(RS!VALOR_TOTAL)
      
   ElseIf optCodigo.Value = True Then
      sSQL = "SELECT cliente.*, os.*, os.status AS var_status FROM cliente INNER JOIN os ON cliente.codigo = os.cod_cliente " & _
         "WHERE (cod_os = " & txtCodigo.Text & ") " & Tipo & " ORDER BY " & INDICE
      
      'SOMAR TODOS =======================
      'Data2.RecordSource = "SELECT SUM(TOTAL)AS VALOR_TOTAL FROM OS WHERE COD_OS = " & txtCodigo.Text & " " & TIPO & ""
      'Data2.Refresh
      'lblTotal.Caption = FormatCurrency(RS!VALOR_TOTAL)
      
   ElseIf optMes.Value = True Then
      sSQL = "SELECT cliente.*, os.*, os.status AS var_status FROM cliente INNER JOIN os ON cliente.codigo = os.cod_cliente " & _
         "WHERE (MONTH(os.data_entrada) = " & cboMES.ListIndex + 1 & ") AND (YEAR(os.data_entrada) = " & cboAno & ") " & Tipo & " ORDER BY " & INDICE
      
      'SOMAR TODOS =======================
      'Data2.RecordSource = "SELECT SUM(TOTAL)AS VALOR_TOTAL FROM OS WHERE Month(OS.DATA_ENTRADA) = " & cboMes.ListIndex + 1 & " AND (Year(OS.DATA_ENTRADA) = " & cboAno & ") " & TIPO & ""
      'Data2.Refresh
      'lblTotal.Caption = FormatCurrency(RS!VALOR_TOTAL)
   
   End If
   
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   
   'Call loadDados
   Montar_Grid_OS r
   lblQtda.Caption = Format(totalRegistros, "00")
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   printSQL = sSQL
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   
   'FORMATAR O GRID
   With Grid
      .Clear
      .Cols = 7
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 0
      .ColWidth(4) = 0
      .ColWidth(5) = 0
      .ColWidth(6) = 0
   End With
   
   Set moCombo = New cComboHelper
End Sub

Private Sub Montar_Grid_OS(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 9
      .Rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 900
      .ColWidth(2) = 700
      .ColWidth(3) = 3100
      .ColWidth(4) = 875
      .ColWidth(5) = 1200
      .ColWidth(6) = 825
      .ColWidth(7) = 825
      .ColWidth(7) = 825
      
      .TextMatrix(0, 1) = "ENT."
      .TextMatrix(0, 2) = "OS"
      .TextMatrix(0, 3) = "CLIENTE"
      .TextMatrix(0, 4) = "MODELO"
      .TextMatrix(0, 5) = "STATUS"
      .TextMatrix(0, 6) = "TOTAL"
      .TextMatrix(0, 7) = "PGTO"
      .TextMatrix(0, 8) = "TIPO"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            For i = 1 To Grid.Rows - 1
               Grid.Row = i
               Grid.Col = 2:   Grid.CellBackColor = vbYellow
               Grid.Col = 7:   Grid.CellBackColor = vbYellow
            Next
            
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("data_entrada"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("cod_os"), "000000")
            .TextMatrix(.Rows - 1, 3) = UCase(rTabela("nome"))
            .TextMatrix(.Rows - 1, 4) = UCase(rTabela("modelo"))
            .TextMatrix(.Rows - 1, 5) = rTabela("var_status")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("total"), ocMONEY)
            .TextMatrix(.Rows - 1, 7) = rTabela("tipo_pagamento")
            .TextMatrix(.Rows - 1, 8) = rTabela("pagamento")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
   'If Grid.Col = 0 Then Exit Sub
   'If Grid.Col = 1 Then
   
   'If Grid.TextMatrix(Grid.Row, 2) = "" Then Exit Sub
   '   Ordem_Servicos_Consulta_Geral_Servicos.txtCodOS.Text = (Grid.TextMatrix(Grid.Row, 2))
   '   Ordem_Servicos_Consulta_Geral_Servicos.Show 1
   'Else
   '   Ordem_Servicos_Consulta_Geral_Servicos.txtCodOS.Text = (Grid.TextMatrix(Grid.Row, 2))
   '   Ordem_Servicos_Consulta_Geral_Servicos.Show 1
   'End If
   
   If Grid.Col = 0 Then Exit Sub
   
   If Grid.Col = 2 Then
      If Grid.TextMatrix(Grid.Row, 2) = "" Then Exit Sub
      
      'Ordem_Servicos_Consulta_Geral_Servicos.txtCodOS.Text = (Grid.TextMatrix(Grid.Row, 2))
      'Ordem_Servicos_Consulta_Geral_Servicos.Show 1
   Else
      'Ordem_Servicos_Consulta_Geral_Parcelas.loadInformacoes CInt(Grid.TextMatrix(Grid.Row, 2))
      'Ordem_Servicos_Consulta_Geral_Parcelas.Show 1
   End If
End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   posX = X
   Label3 = posX
   If Label3.Caption > 0 And Label3.Caption < 149 Then Grid.ToolTipText = ""
   If Label3.Caption > 150 And Label3.Caption < 930 Then Grid.ToolTipText = "Dê um duplo-clique para exibir os itens do Pedido."
   If Label3.Caption > 931 And Label3.Caption < 7230 Then Grid.ToolTipText = ""
   If Label3.Caption > 7231 And Label3.Caption < 8355 Then Grid.ToolTipText = "Dê um duplo-clique para exibir a forma de pgto."
   If Label3.Caption > 8356 And Label3.Caption < 9555 Then Grid.ToolTipText = ""
End Sub

Private Sub optImprimirAprazo_Click()
   cmdLocalizar_Click
End Sub

Private Sub optImprimirAvista_Click()
   cmdLocalizar_Click
End Sub

Private Sub optImprimirTodas_Click()
   cmdLocalizar_Click
End Sub

Private Sub optINDCliente_Click()
   cmdLocalizar_Click
End Sub

Private Sub optINDData_Click()
   cmdLocalizar_Click
End Sub

Private Sub optINDForma_Click()
   cmdLocalizar_Click
End Sub

Private Sub optINDStatus_Click()
   cmdLocalizar_Click
End Sub

Private Sub optINDValor_Click()
   cmdLocalizar_Click
End Sub

Private Sub txtcodigo_GotFocus()
   SelectControl txtCodigo
End Sub

Private Sub mskFim_GotFocus()
   SelectControl mskFim
End Sub

Private Sub mskFim_KeyPress(KeyAscii As Integer)
   mskFim.Mask = "##/##/##"
End Sub

Private Sub mskFim_LostFocus()
   If mskFim.Text = "" Or mskFim.Text = "__/__/__" Then
      mskFim.Mask = ""
      mskFim.Text = ""
      Exit Sub
   Else
      If IsDate(mskFim.Text) Then
         cmdLocalizar.SetFocus
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskFim.SetFocus
         SelectControl mskFim
      End If
   End If
End Sub

Private Sub mskInicio_GotFocus()
   SelectControl mskInicio
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
   mskInicio.Mask = "##/##/##"
End Sub

Sub loadDados(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 7
      .Rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 4300
      .ColWidth(4) = 1000
      .ColWidth(5) = 1100
      .ColWidth(6) = 1100
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "NOME DO CLIENTE"
      .TextMatrix(0, 4) = "VALOR"
      .TextMatrix(0, 5) = "TIPO / PGTO"
      .TextMatrix(0, 6) = "PAGAMENTO"
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            For i = 1 To Grid.Rows - 1
               Grid.Row = i
               Grid.Col = 1:   Grid.CellBackColor = vbYellow
               Grid.Col = 5:   Grid.CellBackColor = vbYellow
            Next
            
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("cod_os"), "000000")
            .TextMatrix(.Rows - 1, 2) = rTabela("data_entrada")
            .TextMatrix(.Rows - 1, 3) = UCase(rTabela("nome"))
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("total"), ocMONEY)
            .TextMatrix(.Rows - 1, 5) = rTabela("tipo_pagamento")
            .TextMatrix(.Rows - 1, 6) = rTabela("pagamento")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub

Sub loadCliente()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT * FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   cboCliente.Clear
   Do While Not r.EOF
      cboCliente.AddItem r("nome")
      cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub mskInicio_LostFocus()
   If mskInicio.Text = "" Or mskInicio.Text = "__/__/__" Then
      mskInicio.Mask = ""
      mskInicio.Text = ""
      Exit Sub
   Else
      If IsDate(mskInicio.Text) Then
         If mskFim.Visible = True Then mskFim.SetFocus
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskInicio.SetFocus
         SelectControl mskInicio
      End If
   End If
End Sub

Private Sub optClientes_Click()
   lblMecanico.Visible = False
   cboMecanico.Visible = False
   
   lblInicio.Visible = False
   mskInicio.Visible = False
   lblFim.Visible = False
   mskFim.Visible = False
   lblAte.Visible = False
   
   lblClientes.Visible = True
   cboCliente.Visible = True
   
   lblCodigo.Visible = False
   txtCodigo.Visible = False
   
   lblMes.Visible = False
   cboMES.Visible = False
   lblAno.Visible = False
   cboAno.Visible = False
   
   Limpar_Objetos
   cboCliente.SetFocus
End Sub

Private Sub optCodigo_Click()
   lblMecanico.Visible = False
   cboMecanico.Visible = False
   
   lblInicio.Visible = False
   mskInicio.Visible = False
   lblFim.Visible = False
   mskFim.Visible = False
   lblAte.Visible = False
   
   lblClientes.Visible = False
   cboCliente.Visible = False
   
   lblCodigo.Visible = True
   txtCodigo.Visible = True
   
   lblMes.Visible = False
   cboMES.Visible = False
   lblAno.Visible = False
   cboAno.Visible = False
   
   Limpar_Objetos
   'txtCodigo.SetFocus
End Sub

Private Sub optDatas_Click()
   lblMecanico.Visible = False
   cboMecanico.Visible = False
   
   lblInicio.Visible = True
   mskInicio.Visible = True
   lblFim.Visible = True
   mskFim.Visible = True
   lblAte.Visible = True
   
   lblClientes.Visible = False
   cboCliente.Visible = False
   
   lblCodigo.Visible = False
   txtCodigo.Visible = False
   
   lblMes.Visible = False
   cboMES.Visible = False
   lblAno.Visible = False
   cboAno.Visible = False
   
   Limpar_Objetos
   mskInicio.SetFocus
End Sub

Private Sub optMes_Click()
   lblMecanico.Visible = False
   cboMecanico.Visible = False
   
   lblInicio.Visible = False
   mskInicio.Visible = False
   lblFim.Visible = False
   mskFim.Visible = False
   lblAte.Visible = False
   
   lblClientes.Visible = False
   cboCliente.Visible = False
   
   lblCodigo.Visible = False
   txtCodigo.Visible = False
   
   lblMes.Visible = True
   cboMES.Visible = True
   lblAno.Visible = True
   cboAno.Visible = True
   
   Limpar_Objetos
   cboMES.SetFocus
End Sub

Private Sub optmecanico_Click()
   lblMecanico.Visible = True
   cboMecanico.Visible = True
   
   lblInicio.Visible = False
   mskInicio.Visible = False
   lblFim.Visible = False
   mskFim.Visible = False
   lblAte.Visible = False
   
   lblClientes.Visible = False
   cboCliente.Visible = False
   
   lblCodigo.Visible = False
   txtCodigo.Visible = False
   
   lblMes.Visible = False
   cboMES.Visible = False
   lblAno.Visible = False
   cboAno.Visible = False
   
   Limpar_Objetos
   cboMecanico.SetFocus
End Sub
