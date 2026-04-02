VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.MDIForm Tela_Principal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ONLINE COMMERCE - Automação Comercial"
   ClientHeight    =   9945
   ClientLeft      =   1065
   ClientTop       =   -855
   ClientWidth     =   16755
   Icon            =   "Tela_Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   840
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   16695
      TabIndex        =   1
      Top             =   0
      Width           =   16755
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   14820
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtSenha 
         Height          =   285
         Left            =   14820
         TabIndex        =   4
         Top             =   300
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtNivel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   16080
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCodFuncionario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   16080
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   1215
      End
      Begin ChamaleonBtn.chameleonButton cmdCaixa 
         Height          =   750
         Left            =   3900
         TabIndex        =   6
         ToolTipText     =   "FINANCEIRO"
         Top             =   10
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":23D2
         PICN            =   "Tela_Principal.frx":23EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdVendasConsRapida 
         Height          =   750
         Left            =   3120
         TabIndex        =   7
         ToolTipText     =   "Consulta Rápida"
         Top             =   10
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":8835
         PICN            =   "Tela_Principal.frx":8851
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConsulta 
         Height          =   750
         Left            =   2340
         TabIndex        =   8
         ToolTipText     =   "Consultas"
         Top             =   10
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":E8AE
         PICN            =   "Tela_Principal.frx":E8CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdOS 
         Height          =   750
         Left            =   10140
         TabIndex        =   9
         ToolTipText     =   "Ordem de Serviços"
         Top             =   10
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":13D8E
         PICN            =   "Tela_Principal.frx":13DAA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConsultaOS 
         Height          =   750
         Left            =   10920
         TabIndex        =   10
         ToolTipText     =   "O.S. Consulta"
         Top             =   10
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":19FC6
         PICN            =   "Tela_Principal.frx":19FE2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCadClientes 
         Height          =   750
         Left            =   780
         TabIndex        =   11
         ToolTipText     =   "Cadastro de Clientes"
         Top             =   10
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":204FD
         PICN            =   "Tela_Principal.frx":20519
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCadProdutos 
         Height          =   750
         Left            =   13260
         TabIndex        =   12
         ToolTipText     =   "Cadastro de Produtos"
         Top             =   0
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":25ED5
         PICN            =   "Tela_Principal.frx":25EF1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdLogon 
         Height          =   750
         Left            =   15
         TabIndex        =   13
         ToolTipText     =   "Backup Nuvem"
         Top             =   10
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":2B8D4
         PICN            =   "Tela_Principal.frx":2B8F0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdContasApagar 
         Height          =   750
         Left            =   5460
         TabIndex        =   14
         ToolTipText     =   "Contas á Pagar"
         Top             =   10
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":2C4E4
         PICN            =   "Tela_Principal.frx":2C500
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdParcelas 
         Height          =   750
         Left            =   6240
         TabIndex        =   15
         ToolTipText     =   "Parcelas"
         Top             =   10
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":2D230
         PICN            =   "Tela_Principal.frx":2D24C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCalculadora 
         Height          =   750
         Left            =   9360
         TabIndex        =   16
         ToolTipText     =   "Calculadora"
         Top             =   10
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":2DE9E
         PICN            =   "Tela_Principal.frx":2DEBA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdResumoFiscal 
         Height          =   750
         Left            =   1560
         TabIndex        =   18
         ToolTipText     =   "Resumo Fiscal"
         Top             =   10
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":2E2B4
         PICN            =   "Tela_Principal.frx":2E2D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdXML 
         Height          =   750
         Left            =   7020
         TabIndex        =   19
         ToolTipText     =   "Entrada de NFe pela XML"
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":2EF63
         PICN            =   "Tela_Principal.frx":2EF7F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdPreco 
         Height          =   750
         Left            =   7800
         TabIndex        =   20
         ToolTipText     =   "Ajuste de estoque"
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":2FADB
         PICN            =   "Tela_Principal.frx":2FAF7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNFe 
         Height          =   750
         Left            =   8580
         TabIndex        =   21
         ToolTipText     =   "Nota Fiscal Eletronica"
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         BTYPE           =   3
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Tela_Principal.frx":304C8
         PICN            =   "Tela_Principal.frx":304E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblAlerta 
         AutoSize        =   -1  'True
         Caption         =   "[ COMPROMISSO ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   14460
         TabIndex        =   17
         Top             =   540
         Visible         =   0   'False
         Width           =   1635
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   9690
      Width           =   16755
      _ExtentX        =   29554
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18336
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "22:00"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Timer trmEstoque2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1020
      Top             =   4920
   End
   Begin VB.Timer trmAlerta 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   1980
      Top             =   4920
   End
   Begin VB.Timer trmAgenda 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1560
      Top             =   4920
   End
   Begin VB.Timer trmEstoque 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   4920
   End
   Begin VB.Menu menusistema 
      Caption         =   "&Sistema"
      Begin VB.Menu Menu_SIS 
         Caption         =   "&Banco de Dados"
         Begin VB.Menu Menu_Sistema_Backup 
            Caption         =   "&Backup"
         End
         Begin VB.Menu Menu_Sistema_LimparTabelas 
            Caption         =   "&Limpar Tabelas"
         End
      End
      Begin VB.Menu Menu_SIS_Config 
         Caption         =   "&Configurações"
      End
      Begin VB.Menu Menu_SIS_Contadores 
         Caption         =   "Contabilidade"
         Begin VB.Menu Menu_SIS_Contador 
            Caption         =   "Contador"
         End
         Begin VB.Menu Menu_SIS_ExportarXML 
            Caption         =   "Exportar XML"
         End
      End
      Begin VB.Menu Menu_CAD_Empresa 
         Caption         =   "&Licença"
      End
      Begin VB.Menu menuseparar1 
         Caption         =   "-"
      End
      Begin VB.Menu menulogoff 
         Caption         =   "&LogOff"
      End
      Begin VB.Menu menusair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu Menu_CAD 
      Caption         =   "&Cadastro"
      Begin VB.Menu Menu_CAD_Clientes 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu Menu_CAD_Fornecedores 
         Caption         =   "&Fornecedores"
      End
      Begin VB.Menu Menu_CAD_Funcionarios 
         Caption         =   "F&uncionários"
      End
      Begin VB.Menu Menu_CAD_Setores 
         Caption         =   "S&etores"
      End
      Begin VB.Menu Menu_CAD_Parcelas 
         Caption         =   "Par&cela"
      End
      Begin VB.Menu Menu_CAD_Usuario 
         Caption         =   "&Usuário"
      End
      Begin VB.Menu Menu_CAD_Transportadora 
         Caption         =   "&Transportadora"
      End
   End
   Begin VB.Menu Menu_Prod 
      Caption         =   "&Produtos"
      Begin VB.Menu Menu_PROD_Cadastro 
         Caption         =   "&Cadastro"
      End
      Begin VB.Menu Menu_PROD_Saida 
         Caption         =   "&Retirada do Estoque Justificada"
      End
      Begin VB.Menu Menu_PROD_Simples 
         Caption         =   "&Ajuste de Estoque"
      End
      Begin VB.Menu Menu_CONS_EstoqueMinimo 
         Caption         =   "E&stoque Minimo"
      End
      Begin VB.Menu Menu_PROD_Compra 
         Caption         =   "Lista de C&ompra"
      End
      Begin VB.Menu Menu_PROD_BuscaRapida 
         Caption         =   "&Busca Rápida"
      End
      Begin VB.Menu Menu_Prod_AjusteTributos 
         Caption         =   "Ajuste de Tributos"
      End
      Begin VB.Menu Menu_Prod_Divisao1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Prod_NotasManifestadas 
         Caption         =   "Notas Manifestadas"
      End
      Begin VB.Menu Menu_PROD_Entrada 
         Caption         =   "&Entrada do Estoque Manual"
      End
      Begin VB.Menu Menu_Entrada_Estoque 
         Caption         =   "&Entrada de Estoque XML"
      End
      Begin VB.Menu Menu_Prod_Divisao2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Prod_Cashback 
         Caption         =   "Cashback"
      End
   End
   Begin VB.Menu MENU_SERV 
      Caption         =   "&Ordem de Serviços"
      Begin VB.Menu MENU_SERV_OS 
         Caption         =   "&Cadastro"
         Enabled         =   0   'False
      End
      Begin VB.Menu MENU_SERV_Consulta 
         Caption         =   "&Consulta"
         Enabled         =   0   'False
      End
      Begin VB.Menu MENU_SERV_Servicos 
         Caption         =   "&Serviços"
      End
      Begin VB.Menu MENU_SERV_Veiculos 
         Caption         =   "&Veículos"
      End
      Begin VB.Menu MENU_SERV_Acessorios 
         Caption         =   "&Acessórios"
      End
      Begin VB.Menu MENU_SERV_Pneus 
         Caption         =   "&Pneus"
      End
   End
   Begin VB.Menu MENU_ALUGUEL 
      Caption         =   "&Aluguel"
      Begin VB.Menu MENU_ALU_Cad_Equipamentos 
         Caption         =   "Cadastro de Equipamentos"
      End
      Begin VB.Menu MENU_ALU_Aluguel 
         Caption         =   "Controle de Aluguel"
      End
   End
   Begin VB.Menu Menu_Fiscal 
      Caption         =   "F&iscal"
      Begin VB.Menu Menu_FAT_NFeCompeta 
         Caption         =   "N&Fe"
      End
      Begin VB.Menu Menu_FAT_NFeOBS 
         Caption         =   "NFe - Observações"
      End
      Begin VB.Menu Menu_FAT_Linha4 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_FAT_MDFe 
         Caption         =   "MDFe"
         Enabled         =   0   'False
      End
      Begin VB.Menu Menu_FAT_Linha1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_FAT_SPED_Fiscal 
         Caption         =   "SPED Fiscal"
      End
      Begin VB.Menu Menu_FAT_Linha2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_FAT_Inventario 
         Caption         =   "Inventário"
      End
      Begin VB.Menu Menu_FAT_Linha3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_FAT_StatusWS 
         Caption         =   "Consulta do Serviço - SEFAZ"
      End
      Begin VB.Menu Menu_FAT_NET 
         Caption         =   "Consulta Conexão de Internet"
      End
   End
   Begin VB.Menu Menu_Contas 
      Caption         =   "&Financeiro"
      Begin VB.Menu Menu_Fin_Parcelas 
         Caption         =   "&Parcelas"
      End
      Begin VB.Menu Menu_Fin_Suprimentos 
         Caption         =   "&Suprimentos"
      End
      Begin VB.Menu Menu_Fin_Sangria 
         Caption         =   "S&angria"
      End
      Begin VB.Menu Menu_Fin_Retirada 
         Caption         =   "Reirada"
      End
      Begin VB.Menu Menu_Fin_Caixa 
         Caption         =   "&Caixa"
      End
      Begin VB.Menu Menu_Fin_traco2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Fin_APagar 
         Caption         =   "À Pagar"
      End
      Begin VB.Menu Menu_Fin_AReceber 
         Caption         =   "Contas Retroativas"
      End
   End
   Begin VB.Menu Menu_FAT 
      Caption         =   "Fatura&mento"
      Begin VB.Menu Menu_CONS_Fluxo 
         Caption         =   "&Fluxo de Caixa"
      End
      Begin VB.Menu Menu_CONS_Lancamentos 
         Caption         =   "&Controle de Saldos"
      End
      Begin VB.Menu Menu_Fin_traco 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Menu_CON 
      Caption         =   "&Consultas"
      Begin VB.Menu Menu_CONS_Vendas 
         Caption         =   "&Vendas"
      End
      Begin VB.Menu Menu_CONS_VendasRecebiveis 
         Caption         =   "Vendas por &Recebíveis"
      End
      Begin VB.Menu Menu_CONS_VendasLucro 
         Caption         =   "Vendas por &Lucro Estimado"
      End
      Begin VB.Menu menuseparar11 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_CONS_VendasPorProdutos 
         Caption         =   "Vendas Por Produtos"
      End
      Begin VB.Menu Menu_CONS_VendasPorProdutosAgrupados 
         Caption         =   "Vendas Por Produtos Agrupados"
      End
      Begin VB.Menu menuseparar7 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_CONS_EntradaPorProdutos 
         Caption         =   "Entradas Por Produtos"
      End
      Begin VB.Menu Menu_CONS_EntradaPorProdAgrupadas 
         Caption         =   "Entradas Por Produtos Agrupados"
      End
      Begin VB.Menu menuseparar10 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_CONS_SaidasPorProdutos 
         Caption         =   "Saídas Por Produtos"
         Enabled         =   0   'False
      End
      Begin VB.Menu Menu_CONS_SaidaPorProdAgrupadas 
         Caption         =   "Saídas Por Produtos Agrupados"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuseparar8 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_CONS_EntradavsSaida 
         Caption         =   "Comparativo - Entradas vs Saídas"
      End
      Begin VB.Menu menuseparar9 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_CONS_Comissoes 
         Caption         =   "Vendedores - Comissões"
      End
      Begin VB.Menu menuseparar12 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_CONS_Parcelas 
         Caption         =   "&Parcelas - Financeiro"
      End
   End
   Begin VB.Menu Menu_IMP 
      Caption         =   "&Impressões"
      Begin VB.Menu Menu_IMP_Etiquetas 
         Caption         =   "&Etiquetas"
      End
      Begin VB.Menu Menu_IMP_Carne 
         Caption         =   "Carnê"
      End
      Begin VB.Menu Menu_IMP_Recibo 
         Caption         =   "&Recibo"
      End
      Begin VB.Menu Menu_IMP_RecAvulso 
         Caption         =   "Recibo A&vulso"
      End
      Begin VB.Menu Menu_IMP_Produtos 
         Caption         =   "&Lista de Produtos"
      End
      Begin VB.Menu Menu_IMP_Aniversariantes 
         Caption         =   "&Aniversariantes"
      End
      Begin VB.Menu Menu_IMP_Clientes 
         Caption         =   "L&ista de Clientes"
      End
   End
   Begin VB.Menu Menu_Diversos 
      Caption         =   "&Diversos"
      Begin VB.Menu Menu_Diversos_Calculadora 
         Caption         =   "&Calculadora"
      End
      Begin VB.Menu Menu_Diversos_Almoxerifado 
         Caption         =   "&Almoxarifado"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Menu_Ajuda 
      Caption         =   "Aj&uda"
      Begin VB.Menu menuajudaonline 
         Caption         =   "&Sobre Nós"
      End
   End
End
Attribute VB_Name = "Tela_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCfg As ConfigItem
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vCodUsuario As Long
Dim vTipoOS As String

Private Sub Conf_TipoEmpresa()
Dim varTipoEmpresa As String
   Set oCfg = sysConfig("TIPO_EMPRESA")
   varTipoEmpresa = oCfg.Value
   
   If varTipoEmpresa < "6" Then
      cmdCadClientes.ToolTipText = "Cadastro de Clientes"
      
      Menu_CAD_Clientes.Caption = "Clientes"
      Menu_Prod.Visible = True
      Menu_CONS_Vendas.Visible = True
      Menu_PROD_BuscaRapida.Visible = True
      Menu_IMP_Produtos.Visible = True

      cmdLogon.Visible = True
      cmdLogon.Left = 10
      cmdCadClientes.Visible = True
      cmdCadClientes.Left = 780
      cmdCadProdutos.Visible = True
      cmdCadProdutos.Left = 1560
      cmdConsulta.Visible = True
      cmdConsulta.Left = 2340
      cmdVendasConsRapida.Visible = True
      cmdVendasConsRapida.Left = 3900
      cmdCaixa.Visible = True
      cmdCaixa.Left = 4680
      cmdContasApagar.Visible = True
      cmdContasApagar.Left = 5460
      cmdParcelas.Visible = True
      cmdParcelas.Left = 6240
      cmdCalculadora.Visible = True
      cmdCalculadora.Left = 9360
      If varTipoEmpresa = "5" Then
        cmdResumoFiscal.Visible = True
        cmdResumoFiscal.Left = 11700
      Else
        cmdResumoFiscal.Visible = True
        cmdResumoFiscal.Left = 10140
      End If
      'cmdOS.Visible = True    'coloquei esses 2 objetos na rotina Habilitar_OS
      'cmdOS.Left = 10140
      'cmdConsultaOS.Visible = True
      'cmdConsultaOS.Left = 10920
      
      'cmdResumoFiscal.Visible = False

   ElseIf varTipoEmpresa = "6" Or varTipoEmpresa = "7" Then
      cmdCadClientes.ToolTipText = "Cadastro de Alunos"
   
      Menu_CAD_Clientes.Caption = "Alunos"
      Menu_Prod.Visible = False
      Menu_CONS_Vendas.Visible = False
      Menu_PROD_BuscaRapida.Visible = False
      Menu_IMP_Produtos.Visible = False
      
      cmdLogon.Visible = True
      cmdLogon.Left = 10
      cmdCadClientes.Visible = True
      cmdCadClientes.Left = 780
      cmdConsulta.Visible = True
      cmdConsulta.Left = 3900
      cmdCaixa.Visible = True
      cmdCaixa.Left = 4680
      cmdContasApagar.Visible = True
      cmdContasApagar.Left = 5460
      cmdParcelas.Visible = True
      cmdParcelas.Left = 6240
      cmdCalculadora.Visible = True
      cmdCalculadora.Left = 9360
      'cmdOS.Visible = True
      'cmdOS.Left = 10140
      'cmdConsultaOS.Visible = True
      'cmdConsultaOS.Left = 10920

      cmdCadProdutos.Visible = False
      cmdOS.Visible = False
      cmdConsultaOS.Visible = False
      cmdVendasConsRapida.Visible = False
      cmdResumoFiscal.Visible = False
      
   ElseIf varTipoEmpresa = "8" Then
      cmdCadClientes.ToolTipText = "Cadastro de Clientes"

      Menu_CAD_Clientes.Caption = "Alunos"
      Menu_Prod.Visible = False
      Menu_CONS_Vendas.Visible = False
      Menu_PROD_BuscaRapida.Visible = False
      Menu_IMP_Produtos.Visible = False

      cmdLogon.Visible = True
      cmdLogon.Left = 10
      cmdCadClientes.Visible = True
      cmdCadClientes.Left = 780
      cmdResumoFiscal.Visible = True
      cmdResumoFiscal.Left = 2340
      cmdConsulta.Visible = True
      cmdConsulta.Left = 3900
      cmdCaixa.Visible = True
      cmdCaixa.Left = 4680
      cmdContasApagar.Visible = True
      cmdContasApagar.Left = 5460
      cmdParcelas.Visible = True
      cmdParcelas.Left = 6240
      cmdCalculadora.Visible = True
      cmdCalculadora.Left = 9360
      If varTipoEmpresa = "5" Then
        cmdResumoFiscal.Visible = True
        cmdResumoFiscal.Left = 11700
      Else
        cmdResumoFiscal.Visible = True
        cmdResumoFiscal.Left = 10140
      End If
      'cmdOS.Visible = True
      'cmdOS.Left = 10140
      'cmdConsultaOS.Visible = True
      'cmdConsultaOS.Left = 10920

      cmdCadProdutos.Visible = False
      cmdOS.Visible = False
      cmdConsultaOS.Visible = False
      cmdVendasConsRapida.Visible = False

      
      cmdCadProdutos.Visible = False
      cmdOS.Visible = False
      cmdConsultaOS.Visible = False
      cmdVendasConsRapida.Visible = False

      cmdResumoFiscal.Visible = True
   End If
   Set oCfg = Nothing
End Sub
Private Sub Habilitar_Aluguel()
Dim oCfg As ConfigItem
Dim bStatus As Boolean

Set oCfg = sysConfig("aluguel")     'Recupera a config deseja
bStatus = CBool(oCfg.Value)         'Converte o valor para booleano
Set oCfg = Nothing                  'Destroi o objeto

'Habilita/desabilida conforme valor
MENU_ALUGUEL.Visible = bStatus

Set oCfg = Nothing
End Sub
Private Sub Habilitar_OS()
Dim oCfg As ConfigItem
Dim bStatus As Boolean

Set oCfg = sysConfig("OS")    'Recupera a config deseja
bStatus = CBool(oCfg.Value)   'Converte o valor para booleano
Set oCfg = Nothing            'Destroi o objeto

'Habilita/desabilida conforme valor
MENU_SERV.Visible = bStatus

cmdConsultaOS.Visible = bStatus
cmdOS.Visible = bStatus

cmdOS.Left = 10140
cmdConsultaOS.Left = 10920

Dim varTipoOS As String
Set oCfg = sysConfig("TIPO_OS")
varTipoOS = oCfg.Value

If varTipoOS = "Automóveis" Then
    MENU_SERV_Veiculos.Visible = True
    MENU_SERV_Acessorios.Visible = True
    MENU_SERV_Pneus.Visible = False
ElseIf varTipoOS = "Motocicletas" Then
    MENU_SERV_Veiculos.Visible = False
    MENU_SERV_Acessorios.Visible = False
    MENU_SERV_Pneus.Visible = False
ElseIf varTipoOS = "Informática" Then
    MENU_SERV_Veiculos.Visible = False
    MENU_SERV_Acessorios.Visible = False
    MENU_SERV_Pneus.Visible = False
ElseIf varTipoOS = "Motores" Then
    MENU_SERV_Veiculos.Visible = False
    MENU_SERV_Acessorios.Visible = False
    MENU_SERV_Pneus.Visible = False
ElseIf varTipoOS = "Gráfica Rápida" Then
    MENU_SERV_Veiculos.Visible = False
    MENU_SERV_Acessorios.Visible = False
    MENU_SERV_Pneus.Visible = False
ElseIf varTipoOS = "Recapadora" Then
    MENU_SERV_Veiculos.Visible = False
    MENU_SERV_Acessorios.Visible = False
    MENU_SERV_Pneus.Visible = True
End If

Set oCfg = Nothing
End Sub
Sub VerificaAgenda()
' On Local Error GoTo TrataErro
'Dim sSQL As String
'Dim r As ADODB.Recordset

'Dim Data As String, Hora As String
'Dim Agenda As String

'Data = Format(Now, ocDATA_EUA)
'Hora = Format(Now, ocHRMN)

'sSQL = "SELECT * FROM compromissos WHERE (data <= '" & Data & "') AND (hora <= '" & Hora & "') AND (status = 'À fazer');"
'Set r = dbData.OpenRecordset(sSQL)

'If r.BOF Or r.RecordCount = 0 Then
'   trmAlerta.Enabled = False
'   lblAlerta.Visible = False
'Else
'   trmAlerta.Enabled = True
'End If

'If r.State <> 0 Then r.Close
'Set r = Nothing
'Exit Sub
   
'TrataErro:
   'ShowMsg "BANCO DE DADOS AUSENTE!" & vbCrLf & "Verifique o cabo da rede.", vbInformation
End Sub

Private Sub cmdCadClientes_Click()
   Clientes_Cadastro.Show 1
End Sub



Private Sub cmdCadProdutos_Click()
   Menu_PROD_Cadastro_Click
End Sub

Private Sub cmdCalculadora_Click()
   Menu_Diversos_Calculadora_Click
End Sub



Private Sub cmdNFe_Click()
Menu_FAT_NFeCompeta_Click
End Sub

Private Sub cmdPreco_Click()
Menu_PROD_Simples_Click
End Sub

Private Sub cmdResumoFiscal_Click()
Notas_Adesivas.Show
End Sub


Private Sub cmdVendasConsRapida_Click()
Produtos_BuscaRapida.Show 1
End Sub

Private Sub cmdXML_Click()
Menu_Entrada_Estoque_Click
End Sub

Private Sub lblalerta_DblClick()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim DATA As String, hora As String
   Dim Agenda As String
   
   DATA = Format(Now, ocDATA)
   hora = Format(Now, ocHRMN)
   
   'sSQL = "SELECT * FROM compromissos WHERE (data <= " & Format$(Data, ocDATA_EUA) & ") AND (hora <= " & Hora & ") AND (status = 'À fazer') ORDER BY hora;"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'If r.BOF Or r.RecordCount = 0 Then
   '   trmAlerta.Enabled = False
   '   lblAlerta.Visible = False
   '   If r.State <> 0 Then r.Close
   '   Set r = Nothing
   '   Exit Sub
   'End If
   
   trmAlerta.Enabled = True
   
   'Do While Not r.EOF
    '  Agenda = vbCrLf & "Código: " & r("codigo") & vbCrLf & _
    '     "Data: " & r("data") & "    Hora: " & r("hora") & vbCrLf & _
    '     "Para: " & r("para") & vbCrLf & _
    '    "Tipo: " & r("tipo") & vbCrLf & _
     '    "Compromisso: " & r("compromisso")
      
    '  ShowMsg Agenda, vbInformation
      
   '   r.MoveNext
  ' Loop
   
   'If r.State <> 0 Then r.Close
  ' Set r = Nothing
End Sub
Private Sub cmdCaixa_Click()
'Principal_Caixa.txtCodFunc.Text = txtCodFuncionario.Text
Principal_Caixa.Show 1
End Sub
Private Sub cmdConsultaOS_Click()
   'MENU_SERV_Consulta_Click
End Sub
Private Sub cmdContasApagar_Click()
   Menu_Fin_APagar_Click
End Sub
Private Sub cmdParcelas_Click()
Menu_Fin_Parcelas_Click
End Sub
Private Sub cmdLogon_Click()
   menulogoff_Click
End Sub
Private Sub cmdOS_Click()
   'MENU_SERV_OS_Click
End Sub
Private Sub MDIForm_Activate()
Habilitar_OS
Habilitar_Aluguel
End Sub
Sub ShowType(ByVal Ctrl As Object)
    'Use the TypeName function to display the class name as text.
    MsgBox (TypeName(Ctrl))
    'Use the TypeOf function to determine the object's type.
    If TypeOf Ctrl Is Button Then
        MsgBox ("The control is a button.")
    ElseIf TypeOf Ctrl Is CheckBox Then
        MsgBox ("The control is a check box.")
    Else
        MsgBox ("The object is some other type of control.")
    End If
End Sub

Private Sub MDIForm_Load()
Dim oCfg As ConfigItem
   
On Local Error Resume Next

If Screen.Height / Screen.TwipsPerPixelX = 768 Then
   Tela_Principal.Picture = LoadPicture(appPathApp & "Tela_Principal_1024x768.jpg")
End If

If Screen.Height / Screen.TwipsPerPixelX = 600 Then
   Tela_Principal.Picture = LoadPicture(appPathApp & "Tela_Principal_800x600.jpg")
End If

'ver a licensa de uso
'sSQL = "SELECT fantasia FROM empresa;"
'Set r = dbData.OpenRecordset(sSQL)

'If Not r.EOF Then
'   StatusBar1.Panels(1).Text = "Este programa está licenciado para " & r("fantasia") & ".  <Denuncia: (89) 9 8817-7036>"
'End If

sSQL = "SELECT codigo, bloqueio, mes_ref, data_bloqueio FROM licenca_pagamentos where pago = 0 order by data_bloqueio;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    Dim vDataBloq As Date
    Dim vDataAtual As Date
    Dim vQuantDia As Integer
    vDataBloq = r("data_bloqueio")
    vDataAtual = Date
    vQuantDia = vDataBloq - vDataAtual
   
    r.MoveFirst
    StatusBar1.Panels(1).Text = "SUA LICENÇA VENCE EM: " & r("data_bloqueio")
    If vQuantDia = 3 Then
        MsgBox "Sua Licença vence em  " & vQuantDia & " dias.", vbInformation, "Aviso do Sistema"
    ElseIf vQuantDia = 2 Then
        MsgBox "Sua Licença vence em  " & vQuantDia & " dias.", vbInformation, "Aviso do Sistema"
    ElseIf vQuantDia = 1 Then
        MsgBox "Sua Licença vence em  " & vQuantDia & " dia.", vbInformation, "Aviso do Sistema"
    ElseIf vQuantDia = 1 Then
        MsgBox "Sua Licença venceu.", vbInformation, "Aviso do Sistema"
    Else
    
    End If
End If

If r.State <> 0 Then r.Close
Set r = Nothing

Habilitar_OS
VerificaAgenda
Conf_TipoEmpresa
StatusBar1.Panels(5).Text = Format(Date, "dd/mm/yy")

Set oCfg = sysConfig("TIPO_OS")
vTipoOS = oCfg.Value
Set oCfg = Nothing

Menu_PROD_Simples.Tag = 3

If vTipoOS = "Automóveis" Then
    MENU_SERV_Veiculos.Visible = True
    MENU_SERV_Acessorios.Visible = True
    MENU_SERV_Pneus.Visible = True
ElseIf vTipoOS = "Motocicletas" Then
    MENU_SERV_Veiculos.Visible = True
    MENU_SERV_Acessorios.Visible = True
    MENU_SERV_Pneus.Visible = True
ElseIf vTipoOS = "Motores" Then
    MENU_SERV_Veiculos.Visible = False
    MENU_SERV_Acessorios.Visible = False
    MENU_SERV_Pneus.Visible = False
ElseIf vTipoOS = "Gráfica Rápida" Then
    MENU_SERV_Veiculos.Visible = False
    MENU_SERV_Acessorios.Visible = False
    MENU_SERV_Pneus.Visible = False
ElseIf vTipoOS = "Informática" Then
    MENU_SERV_Veiculos.Visible = False
    MENU_SERV_Acessorios.Visible = False
    MENU_SERV_Pneus.Visible = False
ElseIf vTipoOS = "Celular" Then
    MENU_SERV_Veiculos.Visible = False
    MENU_SERV_Acessorios.Visible = False
    MENU_SERV_Pneus.Visible = False
ElseIf vTipoOS = "Recapadora" Then
    MENU_SERV_Veiculos.Visible = False
    MENU_SERV_Acessorios.Visible = False
    MENU_SERV_Pneus.Visible = True
ElseIf vTipoOS = "Comunicação Visual" Then
    MENU_SERV_Veiculos.Visible = False
    MENU_SERV_Acessorios.Visible = False
    MENU_SERV_Pneus.Visible = False
ElseIf vTipoOS = "Agrícola" Then
    MENU_SERV_Veiculos.Visible = True
    MENU_SERV_Acessorios.Visible = True
    MENU_SERV_Pneus.Visible = True
End If


End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Set Picture = Nothing
'EncerrarPrograma
KillProcess "OnlineCommerce"
End Sub

Private Sub MENU_ALU_Aluguel_Click()
Aluguel_Cadastro.Hide
Aluguel_Cadastro.txtCodFuncionario.Text = vCodFunc
Aluguel_Cadastro.Show 1
End Sub

Private Sub MENU_ALU_Cad_Equipamentos_Click()
Aluguel_Cadastro_Equipamentos.Show 1
End Sub

Private Sub Menu_CAD_Clientes_Click()
Clientes_Cadastro.Show 1
End Sub

Private Sub Menu_CAD_Empresa_Click()
Empresa_Cadastro.Show
End Sub

Private Sub Menu_CAD_Fornecedores_Click()
Fornecedor_Cadastro.Show 1
End Sub

Private Sub Menu_CAD_Funcionarios_Click()
   If Tela_Principal.txtNivel.Text <> "1" Then
      ShowMsg "Seu nível de acesso não permite a essa operação!", vbInformation
      Exit Sub
   End If
   
   Funcionario_Cadastro.Show 1
End Sub





Private Sub Menu_CAD_Setores_Click()
   Setor_Cadastro.Show 1
End Sub

Private Sub Menu_CAD_Transportadora_Click()
Transportadora.Show 1
End Sub

Private Sub Menu_CAD_Usuario_Click()
Usuario.Show 1
End Sub

Private Sub Menu_CONS_Comissoes_Click()
Funcionario_Comissao.Show 1
End Sub

Private Sub Menu_CONS_EntradaPorProdAgrupadas_Click()
Entrada_Consulta_PorProdutosAgrupadas.Show
End Sub

Private Sub Menu_CONS_EntradaPorProdutos_Click()
Entrada_Consulta_PorProdutos.Show
End Sub

Private Sub Menu_CONS_EntradavsSaida_Click()
Consulta_EntradavsSaida.Show
End Sub

Private Sub Menu_CONS_EstoqueMinimo_Click()
   Consulta_Estoque_Minimo.Show 1
End Sub

Private Sub Menu_CONS_Fluxo_Click()
'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
Load Fluxo_Caixa
Fluxo_Caixa.txtCodFuncionario.Text = Tela_Principal.txtCodFuncionario.Text
Fluxo_Caixa.Show
End Sub

Private Sub Menu_CONS_Lancamentos_Click()
Lanc_Caixa.Show
End Sub

Private Sub Menu_CONS_Parcelas_Click()
Parcelas_Consulta.Show 1
End Sub

Private Sub Menu_CONS_SaidasPorProdutos_Click()
Saida_Consulta_PorProdutos.Show
End Sub

Private Sub Menu_CONS_Vendas_Click()
Vendas_Consulta.Show 1
End Sub

Private Sub Menu_CONS_VendasLucro_Click()
Vendas_Consulta_Lucro.Show 1
'Vendas_Consulta_PorLucro.Show 1
End Sub

Private Sub Menu_CONS_VendasPorProdutos_Click()
Vendas_Consulta_PorProdutos.Show 1
End Sub

Private Sub Menu_CONS_VendasPorProdutosAgrupados_Click()
Vendas_Consulta_PorProdutosAgrupadas.Show
End Sub


Private Sub Menu_CONS_VendasRecebiveis_Click()
Vendas_Consulta_PorRecebiveis.Show 1
End Sub

Private Sub Menu_Diversos_Calculadora_Click()
   On Error GoTo er
   AppActivate "Calculadora", True
   Exit Sub
   
er:
   Select Case Err.Number
      Case 5: Shell "calc.exe", vbNormalFocus
      Case Else: Resume Next
   End Select
End Sub



Private Sub Menu_Entrada_Estoque_Click()
Entrada_Estoque.Show
End Sub

Private Sub Menu_FAT_Inventario_Click()
Inventario_Cadastro.Show 1
End Sub

Private Sub Menu_FAT_MDFe_Click()
'mdfe.Show
End Sub

Private Sub Menu_FAT_NET_Click()
TesteConexaoInternet
End Sub

Private Sub Menu_FAT_NFeCompeta_Click()
NFe_Completa.Show 1
End Sub


Private Sub Menu_FAT_NFeOBS_Click()
NFe_Observacoes.Show 1
End Sub

Private Sub Menu_FAT_StatusWS_Click()
  ConsultaStatus 55
End Sub

Private Sub Menu_Fin_APagar_Click()
'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
Contas_Cadastro.Show 1
End Sub

Public Sub Menu_Fin_AReceber_Click()
Receber_Cadastro.Show 1
End Sub

Public Sub Menu_Fin_Caixa_Click()
'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
Caixa_Controle_semOS.Show
End Sub

Public Sub Menu_Fin_Parcelas_Click()
   'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
'Parcelas.txtCodFuncionario.Text = txtCodFuncionario.Text
Parcelas.Hide
Parcelas.txtCodFuncionario.Text = vCodFunc
Parcelas.Show 1
End Sub

Private Sub Menu_Fin_Retirada_Click()
Caixa_Retirada.Hide
Caixa_Retirada.txtCodFunc.Text = vCodFunc
Caixa_Retirada.Show 1
End Sub

Public Sub Menu_Fin_Sangria_Click()
Caixa_Saida.Hide
Caixa_Saida.txtCodFunc.Text = vCodFunc
Caixa_Saida.Show 1
End Sub

Public Sub Menu_Fin_Suprimentos_Click()
Caixa_Suprimento.Show 1
End Sub

Private Sub Menu_IMP_Aniversariantes_Click()
   Aniversariantes.Show 1
End Sub

Private Sub Menu_IMP_Carne_Click()
Carne.Show 1
End Sub

Private Sub Menu_IMP_Clientes_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

'sSQL = "SELECT *, iif(STATUS = 0 , 'INATIVO' , 'ATIVO') as var_STATUS FROM CLIENTE order by NOME"

sSQL = "SELECT *, CASE status WHEN 0 THEN 'INATIVO' WHEN 1 THEN 'ATIVO' END AS var_status FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

'Unload Me
Set Imp_ListaClientes.Relatorio.Recordset = r
Imp_ListaClientes.Relatorio.Ativar
Unload Imp_ListaClientes

If r.State <> 0 Then r.Close
Set r = Nothing

Me.Show
End Sub

Private Sub Menu_IMP_Etiquetas_Click()
Etiquetas_Impressao.Show 1
End Sub

Private Sub Menu_IMP_Produtos_Click()
 Dim sSQL As String
 Dim r As ADODB.Recordset
 
 sSQL = "SELECT produtos.ref AS var_Ref, produtos.codigo AS varCodProd, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
         "produtos.fabricante AS var_fab, produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.unid_medida AS var_med, produtos.quant_estoque AS var_quant, " & _
         "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda, ((SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) * produtos.quant_estoque) as var_Total " & _
         "FROM produtos " & _
         "WHERE (produtos.ativo = 1) " & _
         "ORDER BY produtos.descricao;"
 
 Set r = dbData.OpenRecordset(sSQL)
 
'Unload Me
Set Imp_ListaProdutos.Relatorio.Recordset = r
Imp_ListaProdutos.Relatorio.Ativar
Unload Imp_ListaProdutos

If Not r Is Nothing Then If r.State <> 0 Then r.Close
Set r = Nothing

Me.Show
End Sub

Private Sub Menu_IMP_RecAvulso_Click()
Recibos_Avulso.Hide
Recibos_Avulso.txtCodFunc.Text = vCodFunc
Recibos_Avulso.Show 1
End Sub

Private Sub Menu_IMP_Recibo_Click()
Recibo.Hide
Recibo.txtCodFunc.Text = vCodFunc
Recibo.Show 1
End Sub


Private Sub Menu_Prod_AjusteTributos_Click()
Produtos_AjusteTributos.Show 1
End Sub

Private Sub Menu_PROD_BuscaRapida_Click()
Produtos_BuscaRapida.Show 1
End Sub

Private Sub Menu_PROD_Cadastro_Click()
   'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
   'Dim oCfg As ConfigItem
   'Dim iVal As Integer
   
   'Set oCfg = sysConfig("PRODUTO")  'Recupera a config desejada
   'iVal = CInt(oCfg.Value)          'Atribui o valor
   'Set oCfg = Nothing               'Destroi o objeto
   
  ' Select Case iVal
      'Case 1
         'Produtos_Cadastro_ComEntrada.Show 1
      'Case 2
         'Produtos_Cadastro_SemEntrada.Show 1
      'Case 3
         Produtos_Cadastro.Show 1
      'Case 4
         'Produtos_Cadastro_Sapataria.Show 1
  ' End Select
End Sub

Private Sub Menu_Prod_Cashback_Click()
Produtos_Cashback.Show 1
End Sub

Private Sub Menu_PROD_Compra_Click()
   Produtos_Comprar.Show 1
End Sub

Private Sub Menu_PROD_Entrada_Click()
Produtos_Entrada.Show
End Sub

Private Sub Menu_Prod_NotasManifestadas_Click()
Notas_Destinadas_Consulta.Show 1
End Sub

Private Sub Menu_PROD_Saida_Click()
Produtos_Saida_Estoque.Show 1
End Sub

Private Sub Menu_PROD_Simples_Click()
Produtos_Estoque_Simples.Hide
Produtos_Estoque_Simples.lblCodUsuario.Caption = vCodFunc
Produtos_Estoque_Simples.Show 1
End Sub

Private Sub MENU_SERV_Acessorios_Click()
'OS_Automoveis_Acessorios.Show 1
End Sub

Private Sub MENU_SERV_Consulta_Click()
   'Ordem_Servicos_Consulta.Show
End Sub

Private Sub MENU_SERV_OS_Click()
'Dim varTipoOS As String
'Set oCfg = sysConfig("TIPO_OS")
'varTipoOS = oCfg.Value

'If varTipoOS = "Automóveis" Then
   'OS_Recapadora.Show 1
'ElseIf varTipoOS = "Motocicletas" Then
'   OS_Recapadora.Show
'ElseIf varTipoOS = "Informática" Then
'   OS_Recapadora.Show
'ElseIf varTipoOS = "Motores" Then
'   OS_Recapadora.Show
'ElseIf varTipoOS = "Gráfica Rápida" Then
'   OS_Recapadora.Show
'ElseIf varTipoOS = "Recapadora" Then
 '  OS_Recapadora.Show
'End If

Set oCfg = Nothing
End Sub

Private Sub MENU_SERV_Pneus_Click()
'OS_Recapadora_Pneus.Show 1
End Sub

Private Sub MENU_SERV_Servicos_Click()
''Dim varTipoOS As String
''Set oCfg = sysConfig("TIPO_OS")
''varTipoOS = oCfg.Value
'If vTipoOS = "Automóveis" Then
'    OS_CAD_Servicos_Geral.Show 1
'ElseIf vTipoOS = "Motocicletas" Then
'    OS_CAD_Servicos_Geral.Show 1
'ElseIf vTipoOS = "Motores" Then
'    OS_CAD_Servicos_Geral.Show 1
'ElseIf vTipoOS = "Gráfica Rápida" Then
'    OS_CAD_Servicos_Geral.Show 1
'ElseIf vTipoOS = "Informática" Then
'    OS_CAD_Servicos_Geral.Show 1
'ElseIf vTipoOS = "Celular" Then
'    OS_CAD_Servicos_Geral.Show 1
'ElseIf vTipoOS = "Recapadora" Then
'    OS_CAD_Servicos_Recapadora.Show 1
'ElseIf vTipoOS = "Comunicação Visual" Then
'    OS_CAD_Servicos_Geral.Show 1
'ElseIf vTipoOS = "Agrícola" Then
'    OS_CAD_Servicos_Geral.Show 1
'End If

'Set oCfg = Nothing
End Sub

Private Sub MENU_SERV_Veiculos_Click()
   'OS_Carros_Cadastro.Show 1
End Sub

Private Sub Menu_SIS_Config_Click()
   Configuracao_Geral.Show 1
End Sub


Private Sub Menu_SIS_Contador_Click()
Contador_Cadastro.Show 1
End Sub

Private Sub Menu_SIS_ExportarXML_Click()
'Exportar_XML.Show 1    'procurar depois pq sumiu
End Sub

Private Sub Menu_Sistema_Backup_Click()
Dim rEmpresa As ADODB.Recordset, xCaminhoBK As String
Dim NomeEmp As String, i As Integer, ComandoSQL As String, e As String, nomeArquivoBK As String

'parte de encontrar o caminho do sistema
sSQL = "SELECT DiretorioXML, razao FROM Empresa"
Set rEmpresa = dbData.OpenRecordset(sSQL)

If Not rEmpresa.EOF Then
    dirXML = IIf(Right(rEmpresa!DiretorioXML, 1) = "\", rEmpresa!DiretorioXML, rEmpresa!DiretorioXML & "\")
End If

xCaminhoBK = dirXML & "backup"

'cria a pasta caso não exista
If Not Existe(xCaminhoBK) Then MkDir xCaminhoBK

If Not Existe(xCaminhoBK) Then Exit Sub

nomeArquivoBK = Format(Date, "yyyy-mm-dd") & "__" & rEmpresa!Razao & ".bak"
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
   iRetorno = CompactarBackup(nomeArquivoBK, xCaminhoBK)
   On Error Resume Next
   Sleep (1000)
   If iRetorno Then
      Kill xCaminhoBK & "\" & nomeArquivoBK
   End If
End If
End Sub

Private Function CompactarBackup(nomeArquivoBK As String, diretorioDestino As String) As Boolean
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

CompactarBackup = True
Exit Function

deuErro:
  CompactarBackup = False
  ''If FormExists("frmECFMsg") Then Unload frmECFMsg
'lblAguarde.Visible = False
End Function

Private Sub Menu_Sistema_LimparTabelas_Click()
If StatusBar1.Panels(2).Text <> "PROGRAMADOR" Then MsgBox "Somento o programador por executar essa tarefa!", vbInformation, "Aviso do Sistema": Exit Sub

If ShowMsg("Deseja limpar todas as tabelas?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

dbData.Execute "DELETE FROM a_pagar ;"
dbData.Execute "DELETE FROM a_pagar_haver ;"
dbData.Execute "DELETE FROM a_receber ;"
dbData.Execute "DELETE FROM a_receber_haver ;"
dbData.Execute "DELETE FROM a_receber_itens ;"
dbData.Execute "DELETE FROM a_receber_parcelas ;"
dbData.Execute "DELETE FROM a_receber_produtos ;"
dbData.Execute "DELETE FROM a_receber_visitas ;"
dbData.Execute "DELETE FROM Aluguel_Cadastro ;"
dbData.Execute "DELETE FROM Aluguel_Cadastro_Equipamento ;"
dbData.Execute "DELETE FROM Aluguel_Cadastro_Itens ;"
dbData.Execute "DELETE FROM caixa_dia ;"
dbData.Execute "DELETE FROM caixa_entrada ;"
dbData.Execute "DELETE FROM caixa_retirada ;"
dbData.Execute "DELETE FROM caixa_saida ;"
dbData.Execute "DELETE FROM caixa_saldo ;"
dbData.Execute "DELETE FROM caixa_saldo_retirada ;"
dbData.Execute "DELETE FROM caixa_troco ;"
dbData.Execute "DELETE FROM carros ;"
dbData.Execute "DELETE FROM cheque ;"
'dbData.Execute "DELETE FROM Cidade ;"
dbData.Execute "DELETE FROM Cliente WHERE codigo >1;"
dbData.Execute "DELETE FROM compromissos ;"
'dbData.Execute "DELETE FROM configuracao ;"
dbData.Execute "DELETE FROM dtproperties ;"
dbData.Execute "DELETE FROM empresa ;"
dbData.Execute "DELETE FROM empresas_desbloueio ;"
dbData.Execute "DELETE FROM EntradaEstoque ;"
dbData.Execute "DELETE FROM EntradaEstoqueItens ;"
dbData.Execute "DELETE FROM entradas ;"
dbData.Execute "DELETE FROM fornecedor ;"
dbData.Execute "DELETE FROM func_permissao ;"
dbData.Execute "DELETE FROM funcionario WHERE codigo > 1;"
dbData.Execute "DELETE FROM licenca_pagamentos ;"
dbData.Execute "DELETE FROM NaturezaOperacaoNF ;"
dbData.Execute "DELETE FROM NFeCartaCorrecao ;"
dbData.Execute "DELETE FROM NFeInutilizacao ;"
dbData.Execute "DELETE FROM NotaFiscal ;"
dbData.Execute "DELETE FROM NotaFiscalAutorizados ;"
dbData.Execute "DELETE FROM NotaFiscalItens ;"
dbData.Execute "DELETE FROM NotaFiscalItensArmamento ;"
dbData.Execute "DELETE FROM NotaFiscalItensCombustivel ;"
dbData.Execute "DELETE FROM NotaFiscalItensMedicamento ;"
dbData.Execute "DELETE FROM NotaFiscalItensVeiculos ;"
dbData.Execute "DELETE FROM NotaFiscalObservacoes ;"
dbData.Execute "DELETE FROM NotaFiscalParcelas ;"
dbData.Execute "DELETE FROM NotaFiscalRecibos ;"
dbData.Execute "DELETE FROM NotaFiscalReferenciada ;"
dbData.Execute "DELETE FROM ObservacoesNFe ;"
dbData.Execute "DELETE FROM OS_Pneus ;"
dbData.Execute "DELETE FROM OS_Servicos_Comunicacao ;"
dbData.Execute "DELETE FROM TbContabilista ;"
dbData.Execute "DELETE FROM TbInventarios ;"
dbData.Execute "DELETE FROM TbInventariosItens ;"
dbData.Execute "DELETE FROM OS ;"
dbData.Execute "DELETE FROM OS_Acessorios ;"
dbData.Execute "DELETE FROM OS_Acessorios_Auto ;"
dbData.Execute "DELETE FROM OS_Equipamento ;"
dbData.Execute "DELETE FROM OS_Equipamento_Auto ;"
'dbData.Execute "DELETE FROM OS_Fabricante_Caminhao ;"
'dbData.Execute "DELETE FROM OS_Fabricante_Moto ;"
'dbData.Execute "DELETE FROM OS_Fabricantes_Carro ;"
'dbData.Execute "DELETE FROM OS_Modelo_Caminhao ;"
'dbData.Execute "DELETE FROM OS_Modelo_Carro ;"
'dbData.Execute "DELETE FROM OS_Modelo_Moto ;"
dbData.Execute "DELETE FROM OS_Servicos ;"
dbData.Execute "DELETE FROM OS_Servicos_Auto ;"
dbData.Execute "DELETE FROM OS_Servicos_recapadora ;"
dbData.Execute "DELETE FROM OS_Situacao ;"
dbData.Execute "DELETE FROM OS_Situacao_Auto ;"
dbData.Execute "DELETE FROM parcelas ;"
dbData.Execute "DELETE FROM parcelas_haver ;"
dbData.Execute "DELETE FROM pedidos ;"
dbData.Execute "DELETE FROM pedidos_itens ;"
dbData.Execute "DELETE FROM Pedidos_Reabertura ;"
dbData.Execute "DELETE FROM Pedidos_Recebedor ;"
dbData.Execute "DELETE FROM produtos WHERE codigo > 1;"
dbData.Execute "DELETE FROM Produtos_Comp ;"
dbData.Execute "DELETE FROM produtos_composicao ;"
dbData.Execute "DELETE FROM produtos_comprar ;"
dbData.Execute "DELETE FROM produtos_entrada WHERE CODIGO > 1;"
dbData.Execute "DELETE FROM produtos_entrada_itens ;"
dbData.Execute "DELETE FROM Produtos_Gas ;"
dbData.Execute "DELETE FROM Produtos_Precos WHERE COD_PRODUTO > 1;"
dbData.Execute "DELETE FROM Produtos_Quant WHERE COD_PRODUTO > 1;"
dbData.Execute "DELETE FROM Produtos_Referencias ;"
dbData.Execute "DELETE FROM produtos_saida ;"
dbData.Execute "DELETE FROM ProdutosFornecedores ;"
dbData.Execute "DELETE FROM recados ;"
dbData.Execute "DELETE FROM setor WHERE COD_SETOR > 1;"
dbData.Execute "DELETE FROM TbNFCe ;"
dbData.Execute "DELETE FROM TbNFCe_Faturas ;"
dbData.Execute "DELETE FROM TbNFCe_Itens ;"
dbData.Execute "DELETE FROM TbNFCe_XML ;"
dbData.Execute "DELETE FROM telefone ;"
dbData.Execute "DELETE FROM Transportadora ;"
dbData.Execute "DELETE FROM Usuario WHERE codigo > 1;"
dbData.Execute "DELETE FROM Usuario_Acessos WHERE Cod_Usuario > 1;"
'dbData.Execute "DELETE FROM Usuario_permissoes ;"
End Sub



Private Sub menuajudaonline_Click()
   Copyright.Show 1
End Sub

Private Sub menulogoff_Click()
'MsgBox "Timer ativo"
Dim DataHora As Date, xCaminhoBK As String
Dim nomeArquivoBK As String
Dim IniciouProcesso As Boolean

   'picAguarde.Visible = False
   DoEvents
   mensagemErro = ""
   iRetorno = False
   
   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
   Set r = dbData.OpenRecordset(sSQL)
        
   If Not r.EOF Then
      dirXML = IIf(Right(r!DiretorioXML, 1) = "\", r!DiretorioXML, r!DiretorioXML & "\")
   Else
      Exit Sub
   End If
        
   xCaminhoBK = dirXML & "backup"
      
   nomeArquivoBK = Retira(r!CNPJ, ".-/ ", UM_A_UM) & ".rar"
        
   DataHora = Now
        
   If Not Existe(xCaminhoBK & "\" & nomeArquivoBK) Then Exit Sub
        
   Me.MousePointer = 11
   iRetorno = GoogleEnviarArquivo(xCaminhoBK & "\" & nomeArquivoBK)
   DoEvents
   If iRetorno Then
      sSQL = "UPDATE empresa SET BackupDataHora = " & FdthrSQL(DataHora)
      SQLExecuta sSQL
   End If
   Me.MousePointer = 0
End Sub

Private Sub menusair_Click()
   EncerrarPrograma
End Sub

Private Sub trmAgenda_Timer()
   VerificaAgenda
End Sub

Private Sub trmAlerta_Timer()
   DoEvents
   lblAlerta.Visible = Not lblAlerta.Visible
End Sub

Private Sub trmEstoque_Timer()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT quant_estoque, quant_min FROM produtos WHERE (quant_estoque < quant_min) ORDER BY codigo;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If r.BOF Then
      StatusBar1.Panels(3).Text = ""
   Else
      trmEstoque2.Enabled = True
   End If
End Sub

Private Sub trmEstoque2_Timer()
   Static bAviso As Boolean
   StatusBar1.Panels(3).Text = IIf(bAviso, "", "ESTOQUE MÍNIMO")
   bAviso = Not bAviso
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
   Select Case Panel.index
      Case 1
      Case 2
      Case 3
         Consulta_Estoque_Minimo.Show 1
      Case 4
      Case 5
   End Select
End Sub

Private Sub txtCodFuncionario_Change()
If txtCodFuncionario.Text = "" Then Exit Sub
vCodUsuario = txtCodFuncionario.Text

If vCodUsuario = False Then
    Menu_CAD_Usuario.Enabled = True
    Menu_CONS_Fluxo.Enabled = False
    Menu_CONS_Lancamentos.Enabled = False
    'Menu_PROD_Cadastro.Enabled = True
    Menu_PROD_Entrada.Enabled = False
    Menu_Entrada_Estoque.Enabled = False
    Menu_PROD_Saida.Enabled = False
    Menu_PROD_Cadastro.Enabled = False
    Menu_PROD_Simples.Enabled = False
    Menu_CAD_Empresa.Enabled = False
    Menu_SIS_Config.Enabled = False
    Menu_Fin_APagar.Enabled = False
    Menu_Sistema_LimparTabelas.Enabled = False
Else
    If LerPermissoesUsuario(vCodUsuario, 3) = True Then
         Menu_PROD_Simples.Enabled = True
     Else
         Menu_PROD_Simples.Enabled = False
    End If
    
    If LerPermissoesUsuario(vCodUsuario, 4) = True Then
         Menu_CAD_Usuario.Enabled = True
     Else
         Menu_CAD_Usuario.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 5) = True Then
         Menu_CONS_Fluxo.Enabled = True
     Else
         Menu_CONS_Fluxo.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 6) = True Then
         Menu_CONS_Lancamentos.Enabled = True
     Else
         Menu_CONS_Lancamentos.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 7) = True Then
         Menu_PROD_Entrada.Enabled = True
     Else
         Menu_PROD_Entrada.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 8) = True Then
         Menu_Entrada_Estoque.Enabled = True
     Else
         Menu_Entrada_Estoque.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 9) = True Then
         Menu_PROD_Saida.Enabled = True
     Else
         Menu_PROD_Saida.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 10) = True Then
         Menu_CAD_Empresa.Enabled = True
     Else
         Menu_CAD_Empresa.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 11) = True Then
         Menu_SIS_Config.Enabled = True
     Else
         Menu_SIS_Config.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 12) = True Then
         Menu_Fin_APagar.Enabled = True
         cmdContasApagar.Enabled = True
     Else
         Menu_Fin_APagar.Enabled = False
         cmdContasApagar.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 13) = True Then
         Menu_Fin_Parcelas.Enabled = True
         'cmdContasApagar.Enabled = True
     Else
         Menu_Fin_Parcelas.Enabled = False
         'cmdContasApagar.Enabled = False

    End If
    If LerPermissoesUsuario(vCodUsuario, 14) = True Then
         Menu_Fin_AReceber.Enabled = True
     Else
         Menu_Fin_AReceber.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 15) = True Then
         Menu_Fin_Sangria.Enabled = True
     Else
         Menu_Fin_Sangria.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 16) = True Then
         Menu_Fin_Retirada.Enabled = True
     Else
         Menu_Fin_Retirada.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 17) = True Then
         Menu_Fin_Suprimentos.Enabled = True
     Else
         Menu_Fin_Suprimentos.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 18) = True Then
         Menu_Fin_Caixa.Enabled = True
     Else
         Menu_Fin_Caixa.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 19) = True Then
         Menu_CONS_Vendas.Enabled = True
     Else
         Menu_CONS_Vendas.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 20) = True Then
         Menu_CONS_VendasPorProdutos.Enabled = True
     Else
         Menu_CONS_VendasPorProdutos.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 21) = True Then
         Menu_CONS_VendasPorProdutosAgrupados.Enabled = True
     Else
         Menu_CONS_VendasPorProdutosAgrupados.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 22) = True Then
         Menu_CONS_EntradaPorProdutos.Enabled = True
     Else
         Menu_CONS_EntradaPorProdutos.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 23) = True Then
         Menu_CONS_EntradaPorProdAgrupadas.Enabled = True
     Else
         Menu_CONS_EntradaPorProdAgrupadas.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 24) = True Then
         Menu_CONS_EntradavsSaida.Enabled = True
     Else
         Menu_CONS_EntradavsSaida.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 25) = True Then
         Menu_CONS_Comissoes.Enabled = True
     Else
         Menu_CONS_Comissoes.Enabled = False
    End If
    If LerPermissoesUsuario(vCodUsuario, 26) = True Then
         Menu_CONS_Parcelas.Enabled = True
     Else
         Menu_CONS_Parcelas.Enabled = False
    End If

    If LerPermissoesUsuario(vCodUsuario, 27) = True Then
         Menu_PROD_Cadastro.Enabled = True
     Else
         Menu_PROD_Cadastro.Enabled = False
    End If

    If LerPermissoesUsuario(vCodUsuario, 29) = True Then
         Menu_CONS_VendasRecebiveis = True
     Else
         Menu_CONS_VendasRecebiveis = False
    End If
    
    If LerPermissoesUsuario(vCodUsuario, 30) = True Then
         Menu_CONS_VendasLucro = True
     Else
         Menu_CONS_VendasLucro = False
    End If
End If
End Sub
Public Function LerPermissoesUsuario(vCodUser As Long, permissao As Long) As Boolean
sSQL = "SELECT Usuario_Acessos.Cod_Permissao FROM Usuario INNER JOIN Usuario_Acessos ON Usuario.Codigo = Usuario_Acessos.Cod_Usuario WHERE (Usuario_Acessos.Cod_Usuario = " & vCodUser & ") AND Usuario_Acessos.Cod_Permissao = " & permissao & ";"
Set r = dbData.OpenRecordset(sSQL)

If r.EOF And r.BOF Then
   LerPermissoesUsuario = False ' não achou a permissao
Else
   LerPermissoesUsuario = True 'aqui achou
End If
End Function

