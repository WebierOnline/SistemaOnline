VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Caixa_Suprimento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SUPRIMENTO"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "Caixa_Suprimento.frx":0000
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   8865
      TabIndex        =   26
      Top             =   0
      Width           =   8895
      Begin VB.Image Image2 
         Height          =   555
         Left            =   180
         Picture         =   "Caixa_Suprimento.frx":23D2
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SUPRIMENTO"
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
         Left            =   3300
         TabIndex        =   27
         Top             =   300
         Width           =   2040
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   60
      TabIndex        =   14
      Top             =   1020
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CADASTRO"
      TabPicture(0)   =   "Caixa_Suprimento.frx":2BBD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdNovo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdExcluir"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAlterar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSalvar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSair"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Picture1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Caixa_Suprimento.frx":2BD9
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblQuant"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblValor"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frmConsulta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.PictureBox Picture1 
         Height          =   4155
         Left            =   -74880
         ScaleHeight     =   4095
         ScaleWidth      =   6555
         TabIndex        =   19
         Top             =   480
         Width           =   6615
         Begin VB.Frame frmCadastro 
            Height          =   1635
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   6435
            Begin VB.ComboBox cboForma 
               Height          =   315
               Left            =   2280
               TabIndex        =   2
               Top             =   1200
               Width           =   1575
            End
            Begin VB.TextBox txtCodigo 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   180
               TabIndex        =   25
               Top             =   0
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.ComboBox cboDesc 
               Height          =   315
               Left            =   120
               TabIndex        =   0
               Top             =   480
               Width           =   6195
            End
            Begin VB.TextBox txtValor 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5160
               TabIndex        =   4
               Top             =   1200
               Width           =   1155
            End
            Begin VB.ComboBox cboSetor 
               Height          =   315
               ItemData        =   "Caixa_Suprimento.frx":2BF5
               Left            =   120
               List            =   "Caixa_Suprimento.frx":2BF7
               TabIndex        =   1
               Top             =   1200
               Width           =   2115
            End
            Begin MSMask.MaskEdBox mskData 
               Height          =   315
               Left            =   3900
               TabIndex        =   3
               Top             =   1200
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label lblFormaPgto 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H80000009&
               BackStyle       =   0  'Transparent
               Caption         =   "Forma de Pgto"
               Height          =   195
               Left            =   2325
               TabIndex        =   34
               Top             =   960
               Width           =   1035
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição"
               Height          =   195
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data"
               Height          =   195
               Left            =   3900
               TabIndex        =   23
               Top             =   960
               Width           =   345
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   5100
               TabIndex        =   22
               Top             =   960
               Width           =   360
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Setor"
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   960
               Width           =   375
            End
         End
      End
      Begin VB.Frame frmConsulta 
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
         Height          =   4095
         Left            =   120
         TabIndex        =   15
         Top             =   420
         Width           =   8595
         Begin VB.OptionButton optConsMes 
            Caption         =   "&Mês"
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
            Left            =   1920
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   6300
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   180
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboMES 
            Height          =   315
            ItemData        =   "Caixa_Suprimento.frx":2BF9
            Left            =   4500
            List            =   "Caixa_Suprimento.frx":2BFB
            TabIndex        =   9
            Top             =   180
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.OptionButton optTodos 
            Caption         =   "&Todos"
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
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optConsData 
            Caption         =   "&Data"
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
            Left            =   1080
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.CommandButton cmdExibir 
            Caption         =   "&Exibir"
            Height          =   315
            Left            =   7680
            TabIndex        =   11
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSMask.MaskEdBox mskConsData 
            Height          =   315
            Left            =   6420
            TabIndex        =   12
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridSuprimentos 
            Height          =   3435
            Left            =   60
            TabIndex        =   33
            Top             =   540
            Width           =   8475
            _ExtentX        =   14949
            _ExtentY        =   6059
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.Label lblCONmes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E&scolha o mês/ano:"
            Height          =   195
            Left            =   3060
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   1425
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdSair 
         Height          =   615
         Left            =   -68100
         TabIndex        =   5
         Top             =   3780
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "Caixa_Suprimento.frx":2BFD
         PICN            =   "Caixa_Suprimento.frx":2C19
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvar 
         Height          =   615
         Left            =   -68100
         TabIndex        =   28
         Top             =   1140
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "Caixa_Suprimento.frx":2F33
         PICN            =   "Caixa_Suprimento.frx":2F4F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterar 
         Height          =   615
         Left            =   -68100
         TabIndex        =   29
         Top             =   2460
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Alterar"
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
         MICON           =   "Caixa_Suprimento.frx":9819
         PICN            =   "Caixa_Suprimento.frx":9835
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   -68100
         TabIndex        =   30
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "Caixa_Suprimento.frx":A10F
         PICN            =   "Caixa_Suprimento.frx":A12B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExcluir 
         Height          =   615
         Left            =   -68100
         TabIndex        =   31
         Top             =   3120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Excluir"
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
         MICON           =   "Caixa_Suprimento.frx":10BCF
         PICN            =   "Caixa_Suprimento.frx":10BEB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNovo 
         Height          =   615
         Left            =   -68100
         TabIndex        =   32
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Novo"
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
         MICON           =   "Caixa_Suprimento.frx":10F05
         PICN            =   "Caixa_Suprimento.frx":10F21
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblValor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "R$ 0,00"
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
         Height          =   195
         Left            =   7995
         TabIndex        =   18
         Top             =   4530
         Width           =   690
      End
      Begin VB.Label lblQuant 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   4530
         Width           =   225
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   8520
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   35
      Top             =   5910
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9393
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2470
            MinWidth        =   2470
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
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
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "00:00"
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
      Left            =   7980
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Caixa_Suprimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Dim ULTRAPASSOU_VALOR As Boolean
Dim CAIXA_FECHADO As Boolean

Private Function Atualizar_Dados() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE caixa_entrada SET " & _
      "descricao = '" & cboDesc.Text & "', " & _
      "setor = '" & cboSetor.Text & "', " & _
      "forma_pgto = '" & cboForma.Text & "', " & _
      "data = CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & _
      "valor = " & Replace(CCur(txtValor.Text), ",", ".") & ", " & _
      "hora = '" & Format$(lblHora, ocHRMN) & "' "
   
   'Condição para atualização
   sSQL = sSQL & "WHERE (codigo = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Function Inserir_Dados(ByVal Codigo As Long) As Boolean
   'A inclusão deve ser feita utilizando o comando INSERT INTO do sql
   'e não mais usando o método .AddNew do Recordset
   
   Dim sSQL As String
   
   'Comando de inclusão
   sSQL = "INSERT INTO caixa_entrada (codigo, descricao, setor, forma_pgto, data, valor, hora, maquina) VALUES (" & _
      Codigo & ", '" & cboDesc.Text & "', '" & cboSetor.Text & "', '" & cboForma.Text & "', CONVERT(DATETIME, '" & _
      Format$(mskData.Text, ocDATA) & "', 103), " & Replace(CCur(txtValor.Text), ",", ".") & ", '" & StatusBar1.Panels(3).Text & "','" & _
      StatusBar1.Panels(2).Text & "');"
   
   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Autonumeracao() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 0
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_entrada FROM caixa_entrada;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("cod_entrada") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   Autonumeracao = lRet
End Function

Private Sub Campos_Brancos()
   txtCodigo.Text = ""
   mskData.Mask = ""
   mskData.Text = ""
   txtValor.Text = ""
   cboDesc.Text = ""
   cboDesc.Clear
   cboForma.Text = ""
   cboSetor.Text = ""
   cboSetor.Clear
End Sub

Private Sub Criar_Novo()
   Campos_Brancos
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
   '
   'iAno = iAno + 1
   'For i = iAno To LastYear
   '   cboAno.AddItem
   'Next
End Sub

Private Sub cboAno_LostFocus()
   If cboAno.Text = "" Then Exit Sub Else cmdExibir.SetFocus
End Sub

Private Sub cboForma_GotFocus()
   cboForma.Clear
   cboForma.AddItem UCase("DINHEIRO")
   cboForma.AddItem UCase("CHEQUE")
   cboForma.AddItem UCase("CARTAO")
   cboForma.AddItem UCase("DEPOSITO")
   cboForma.AddItem UCase("TRANSFERENCIA")
   
   If cboForma.ListCount <> 0 Then cboForma.ListIndex = 0
   
   moCombo.AttachTo cboForma
End Sub

Private Sub cboMes_GotFocus()
   Dim vMes As Integer
   
   cboMES.Clear
   
   For vMes = 1 To 12
      cboMES.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboMES
End Sub

Private Sub cboMes_LostFocus()
   If cboMES.Text = "" Then Exit Sub Else cboAno.SetFocus
End Sub

Private Sub cboSETOR_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdAlterar_Click()
'VERIFICAR O STATUS DO CAIXA
Dim cStatus As Integer
cStatus = Verificar_Caixa
Select Case cStatus
   Case -1
      ShowMsg "Este caixa ainda não foi aberto.", vbExclamation
      Exit Sub
   Case 1
      ShowMsg "O caixa está fechado!", vbExclamation
      Exit Sub
End Select

   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodigo.Text = "" Then Exit Sub
   
   'Faz a atualização de forma direta e verifica se houve algum erro
   If Not Atualizar_Dados Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Campos_Brancos
   Form_Load
End Sub

Private Sub cmdCancelar_Click()
   Campos_Brancos
   Form_Load
End Sub

Private Sub cmdExcluir_Click()
'VERIFICAR O STATUS DO CAIXA
Dim cStatus As Integer
cStatus = Verificar_Caixa
Select Case cStatus
   Case -1
      ShowMsg "Este caixa ainda não foi aberto.", vbExclamation
      Exit Sub
   Case 1
      ShowMsg "O caixa está fechado!", vbExclamation
      Exit Sub
End Select

   Dim sSQL As String
   Dim bRet As Boolean
   
   If txtCodigo.Text = "" Then Exit Sub
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Excluir esse Suprimento?", vbInformation + vbYesNo) = vbNo Then Exit Sub

   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM caixa_entrada WHERE (codigo = " & txtCodigo.Text & ");"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   Campos_Brancos
   Form_Load
End Sub

Private Sub cmdExibir_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   If optTodos.Value = True Then
      sSQL = "SELECT * FROM caixa_entrada ORDER BY data, hora;"
      
   ElseIf optConsData.Value = True Then
      If Not IsDate(mskConsData) Then Exit Sub
      sSQL = "SELECT * FROM caixa_entrada WHERE (data = CONVERT(DATETIME, '" & Format(mskConsData.Text, ocDATA) & "', 103)) ORDER BY data, hora;"
      
   ElseIf optConsMes.Value = True Then
      If cboMES.Text = "" Or cboMES.ListIndex = -1 Then Exit Sub
      If cboAno.Text = "" Or cboAno.ListIndex = -1 Then Exit Sub
      sSQL = "SELECT * FROM caixa_entrada WHERE (MONTH(data) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data) = " & cboAno & ") ORDER BY data, hora;"
       
   End If
   
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   FormatarGrid r
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'MOSTRAR A QUANTIDADE REGISTROS
   lblQuant.Caption = Format(totalRegistros, "00")
End Sub

Private Sub cmdNovo_Click()
'VERIFICAR O STATUS DO CAIXA
Dim cStatus As Integer
cStatus = Verificar_Caixa
Select Case cStatus
   Case -1
      ShowMsg "Este caixa ainda não foi aberto.", vbExclamation
      Exit Sub
   Case 1
      ShowMsg "O caixa está fechado!", vbExclamation
      Exit Sub
End Select
   
   frmCadastro.Enabled = True
   Criar_Novo
   cmdNovo.Enabled = False
   cmdSalvar.Enabled = True
   cmdCancelar.Enabled = True
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   cboDesc.SetFocus
End Sub

Private Function Verificar_Caixa() As Integer
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim cxaStatus As Integer
  
   cxaStatus = -1   'Não foi aberto
   If cmdAlterar.Enabled = True Then
      sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(mskData.FormattedText, ocDATA) & "', 103)) AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
   Else
      sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(StatusBar1.Panels(4), ocDATA) & "', 103)) AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
   End If
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then cxaStatus = CInt(ValidateNull(r("status")))   '0 = aberto, 1 = fechado
   If r.State <> 0 Then r.Close
   Set r = Nothing
   Verificar_Caixa = cxaStatus
End Function

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cboDesc_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboDesc.Clear
   
   sSQL = "SELECT DISTINCT descricao FROM caixa_entrada ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDesc.AddItem r("descricao")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboDesc
End Sub

Private Sub cboDesc_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSetor_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboSetor.Clear
   
   sSQL = "SELECT DISTINCT setor FROM setor ORDER BY setor;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboSetor.AddItem r("setor")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   If cboSetor.ListCount <> 0 Then cboSetor.ListIndex = 0
   moCombo.AttachTo cboSetor
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lNovoCod As Long
   
   If txtValor.Text = "" Or cboDesc.Text = "" Or cboForma.Text = "" Then
      ShowMsg "Formulário incompleto!", vbInformation
      cboDesc.SetFocus
      Exit Sub
   End If
   
   'ADICIONAR REGISTRO
   lNovoCod = Autonumeracao
   
   'Faz a inserção de forma direta e verifica se houve algum erro
   If Not Inserir_Dados(lNovoCod) Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Campos_Brancos
   Form_Load
   Exit Sub
   
TrataErro:
   If Err.Number = 3022 Then
      ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
   SSTab1.Tab = 0
   cmdNovo.Enabled = True
   frmCadastro.Enabled = False
   cmdNovo.Enabled = True
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   
   optConsData_Click
   cmdExibir_Click
   
   'colocar o nome da maquina na barra de status
   Dim var_Maquina As String
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
   Set oIni = Nothing
   
   StatusBar1.Panels(2).Text = var_Maquina
   StatusBar1.Panels(4).Text = Format(Date, "dd/mm/yy")
   StatusBar1.Panels(3).Text = Format(Now, "hh:mm")
   
   Set moCombo = New cComboHelper
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With GridSuprimentos
      .Clear
      .Cols = 7
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 900
      .ColWidth(3) = 800
      .ColWidth(4) = 2200
      .ColWidth(5) = 1300
      .ColWidth(6) = 1000
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "HORA"
      .TextMatrix(0, 4) = "DESC"
      .TextMatrix(0, 5) = "SETOR"
      .TextMatrix(0, 6) = "VALOR"
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("data"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("hora"), ocHRMN)
            .TextMatrix(.Rows - 1, 4) = rTabela("descricao")
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("setor"))
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("valor"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 6
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   lblValor.Caption = Format(SomaGrid(GridSuprimentos, 6), ocMONEY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To var_Grid.Rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CCur(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub GridSuprimentos_DblClick()
   frmCadastro.Enabled = True
   cmdNovo.Enabled = True
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = True
   cmdExcluir.Enabled = True
   txtCodigo.Text = ""
   txtCodigo.Text = (GridSuprimentos.TextMatrix(GridSuprimentos.Row, 1))
End Sub

Private Sub mskConsData_GotFocus()
   SelectControl mskConsData
End Sub

Private Sub mskConsData_KeyPress(KeyAscii As Integer)
   mskConsData.Mask = "##/##/##"
End Sub

Private Sub mskConsData_LostFocus()
   If mskConsData.Text = "" Or mskConsData.Text = "__/__/__" Then
      mskConsData.Mask = ""
      mskConsData.Text = ""
   Else
      If IsDate(mskConsData.Text) Then
         Exit Sub
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskConsData.SetFocus
         SelectControl mskConsData
      End If
   End If
End Sub

Private Sub mskData_GotFocus()
   mskData.Text = Format(Date, "dd/mm/yy")
   SelectControl mskData
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
   mskData.Mask = "##/##/##"
End Sub

Private Sub mskData_LostFocus()
   If mskData.Text = "" Or mskData.Text = "__/__/__" Then
      mskData.Mask = ""
      mskData.Text = ""
   Else
      If IsDate(mskData.Text) Then
         Exit Sub
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskData.SetFocus
         SelectControl mskData
      End If
   End If
End Sub

Private Sub optConsData_Click()
   mskConsData.Visible = True
   cboMES.Visible = False
   cboAno.Visible = False
   lblCONmes.Visible = False
   cmdExibir.Visible = True
   mskConsData.Text = Format(Date, "dd/mm/yy")
   If SSTab1.Tab = 1 Then mskConsData.SetFocus
End Sub

Private Sub optConsMes_Click()
   mskConsData.Visible = False
   cboMES.Visible = True
   cboAno.Visible = True
   cmdExibir.Visible = True
   lblCONmes.Visible = True
   cboMES.SetFocus
End Sub

Private Sub optTodos_Click()
   mskConsData.Visible = False
   cboMES.Visible = False
   cboAno.Visible = False
   lblCONmes.Visible = False
   cmdExibir_Click
   cmdExibir.Visible = False
End Sub

Private Sub Timer1_Timer()
   lblHora.Caption = Format(Time, ocHRMN)
End Sub

Private Sub txtCodigo_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodigo.Text = "" Then Exit Sub
   
   If cmdAlterar.Visible = True Then
      sSQL = "SELECT * FROM caixa_entrada WHERE (codigo = " & txtCodigo.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then Mostrar_Entrada r
      If r.State <> 0 Then r.Close
      Set r = Nothing
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Mostrar_Entrada(rTabela As ADODB.Recordset)
   If Not rTabela Is Nothing Then
      cboDesc.Text = rTabela("descricao")
      cboSetor.Text = rTabela("setor")
      cboForma.Text = rTabela("forma_pgto")
      mskData.Text = Format(rTabela("data"), "dd/mm/yy")
      txtValor.Text = Format(rTabela("valor"), ocMONEY)
   End If
End Sub

Private Sub txtValor_GotFocus()
   SelectControl txtValor
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtValor_LostFocus()
   If txtValor.Text = "" Then
      txtValor.Text = Format(0, ocMONEY)
   Else
      txtValor.Text = Format(txtValor, ocMONEY)
   End If
End Sub
