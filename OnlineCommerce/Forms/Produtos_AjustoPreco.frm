VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form Produtos_AjustoPreco 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AJUSTE DE PREÇOS"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmProduto 
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   60
      TabIndex        =   27
      Top             =   60
      Width           =   11175
      Begin VB.Frame Frame5 
         Caption         =   "ÚLTIMO PREÇO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1395
         Left            =   120
         TabIndex        =   36
         Top             =   900
         Width           =   10935
         Begin VB.Frame Frame10 
            Caption         =   "Custo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   180
            TabIndex        =   57
            Top             =   300
            Width           =   1695
            Begin VB.TextBox txtCustoAnt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   60
               Locked          =   -1  'True
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   480
               Width           =   1335
            End
            Begin ChamaleonBtn.chameleonButton cmdRepetirLucro 
               Height          =   315
               Left            =   1440
               TabIndex        =   61
               Top             =   480
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   ">"
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
               MICON           =   "Produtos_AjustoPreco.frx":0000
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Último Vlr Custo"
               Height          =   195
               Left            =   120
               TabIndex        =   59
               Top             =   240
               Width           =   1110
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Atacado - À Prazo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   8580
            TabIndex        =   52
            Top             =   300
            Width           =   2175
            Begin VB.TextBox txtMargemAPant 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   480
               Width           =   895
            End
            Begin VB.TextBox txtValorAPant 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1080
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   180
               TabIndex        =   56
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1140
               TabIndex        =   55
               Top             =   240
               Width           =   360
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Atacado - À Vista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   6360
            TabIndex        =   47
            Top             =   300
            Width           =   2175
            Begin VB.TextBox txtValorAVant 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1080
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtMargemAVant 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   480
               Width           =   895
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1140
               TabIndex        =   51
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   180
               TabIndex        =   50
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Varejo - À Prazo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   4140
            TabIndex        =   42
            Top             =   300
            Width           =   2175
            Begin VB.TextBox txtMargemVPant 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   480
               Width           =   895
            End
            Begin VB.TextBox txtValorVPant 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1080
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   180
               TabIndex        =   46
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1140
               TabIndex        =   45
               Top             =   240
               Width           =   360
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Varejo - À vista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   1920
            TabIndex        =   37
            Top             =   300
            Width           =   2175
            Begin VB.TextBox txtValorVVant 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1080
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtMargemVVAnt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   480
               Width           =   895
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1140
               TabIndex        =   41
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   180
               TabIndex        =   40
               Top             =   240
               Width           =   735
            End
         End
      End
      Begin VB.TextBox txtQuant 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtFabricante 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   480
         Width           =   1995
      End
      Begin VB.TextBox txtCodProduto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   90
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quant. Atual"
         Height          =   195
         Left            =   10020
         TabIndex        =   35
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante"
         Height          =   195
         Left            =   7980
         TabIndex        =   33
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód."
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1320
         TabIndex        =   29
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame frmPrecos 
      Caption         =   "Preços"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   60
      TabIndex        =   11
      Top             =   2580
      Width           =   11175
      Begin VB.Frame Frame1 
         Caption         =   "Varejo - À vista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   1860
         TabIndex        =   23
         Top             =   300
         Width           =   2415
         Begin VB.TextBox txtMargemVV 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   895
         End
         Begin VB.TextBox txtValorVV 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   2
            Top             =   480
            Width           =   975
         End
         Begin ChamaleonBtn.chameleonButton cmdRepetir 
            Height          =   315
            Left            =   2100
            TabIndex        =   60
            Top             =   480
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   ">"
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
            MICON           =   "Produtos_AjustoPreco.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margem %"
            Height          =   195
            Left            =   180
            TabIndex        =   25
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   195
            Left            =   1140
            TabIndex        =   24
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Varejo - À Prazo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   4320
         TabIndex        =   20
         Top             =   300
         Width           =   2175
         Begin VB.TextBox txtValorVP 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   4
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtMargemVP 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   895
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   195
            Left            =   1140
            TabIndex        =   22
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margem %"
            Height          =   195
            Left            =   180
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Atacado - À Vista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   6540
         TabIndex        =   17
         Top             =   300
         Width           =   2175
         Begin VB.TextBox txtMargemAV 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   895
         End
         Begin VB.TextBox txtValorAV 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margem %"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   195
            Left            =   1140
            TabIndex        =   18
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Atacado - À Prazo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   8760
         TabIndex        =   14
         Top             =   300
         Width           =   2175
         Begin VB.TextBox txtValorAP 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   8
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtMargemAP 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   180
            TabIndex        =   7
            Top             =   480
            Width           =   895
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   195
            Left            =   1140
            TabIndex        =   16
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margem %"
            Height          =   195
            Left            =   180
            TabIndex        =   15
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame frmCusto 
         Caption         =   "Custo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   1695
         Begin VB.TextBox txtCusto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   60
            TabIndex        =   0
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Último Vlr Custo"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1110
         End
      End
      Begin VB.Label lblAviso 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pressione [ F2 ]  para obter o lucro estimado"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1980
         TabIndex        =   26
         Top             =   1200
         Visible         =   0   'False
         Width           =   3300
         WordWrap        =   -1  'True
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdSalvar 
      Height          =   615
      Left            =   60
      TabIndex        =   9
      Top             =   4260
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Salvar"
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
      MICON           =   "Produtos_AjustoPreco.frx":0038
      PICN            =   "Produtos_AjustoPreco.frx":0054
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
      Left            =   9060
      TabIndex        =   10
      Top             =   4260
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Produtos_AjustoPreco.frx":1DE6
      PICN            =   "Produtos_AjustoPreco.frx":1E02
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
Attribute VB_Name = "Produtos_AjustoPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String
Dim r As ADODB.Recordset
Private Sub LimparObjetos_Produtos()
txtDescricao.Text = ""
txtFabricante.Text = ""
txtQuant.Text = ""
txtMargemVV.Text = ""
txtMargemVP.Text = ""
txtMargemAV.Text = ""
txtMargemAP.Text = ""
txtValorVV.Text = ""
txtValorVP.Text = ""
txtValorAV.Text = ""
txtValorAP.Text = ""
txtCusto.Text = ""
txtMargemVVAnt.Text = ""
txtMargemVPant.Text = ""
txtMargemAVant.Text = ""
txtMargemAPant.Text = ""
txtValorVVant.Text = ""
txtValorVPant.Text = ""
txtValorAVant.Text = ""
txtValorAPant.Text = ""
txtCustoAnt.Text = ""
End Sub
Private Sub MostrarDados_Produto(rTabela As ADODB.Recordset)
txtDescricao.Text = ValidateNull(rTabela("descricao"))
txtFabricante.Text = ValidateNull(rTabela("fabricante"))
txtQuant.Text = ValidateNull(rTabela("quant_estoque"))
End Sub

Private Sub MostrarObjetosPrecos()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodProduto.Text = "" Then Exit Sub

sSQL = "SELECT TOP 1 * FROM Produtos_Precos WHERE (COD_PRODUTO = " & txtCodProduto.Text & ") order by CODIGO desc;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    txtCustoAnt.Text = Format$(r("CUSTO"), ocMONEY)
    txtValorVVant.Text = Format$(r("VALOR_VV"), ocMONEY)
    txtValorVPant.Text = Format$(r("VALOR_VP"), ocMONEY)
    txtValorAVant.Text = Format$(r("VALOR_AV"), ocMONEY)
    txtValorAPant.Text = Format$(r("VALOR_AP"), ocMONEY)
    txtMargemVVAnt.Text = FormatNumber(r("MARGEM_VV"), 2) & "%"
    txtMargemVPant.Text = FormatNumber(r("MARGEM_VP"), 2) & "%"
    txtMargemAVant.Text = FormatNumber(r("MARGEM_AV"), 2) & "%"
    txtMargemAPant.Text = FormatNumber(r("MARGEM_AP"), 2) & "%"
End If
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub



Private Sub cmdRepetir_Click()
txtMargemVP.Text = txtMargemVV.Text
txtMargemAV = txtMargemVV.Text
txtMargemAP = txtMargemVV.Text
CalcularPrecos
End Sub


Private Sub cmdRepetirLucro_Click()
txtCusto.Text = txtCustoAnt.Text
End Sub

Private Sub cmdSair_Click()
Produtos_AjustoPreco.Hide
Produtos_Estoque_Simples.Show 1
End Sub

Private Sub cmdSalvar_Click()
Dim AutoNumeracao As Long

If txtMargemVV.Text = "" Then Exit Sub
If txtMargemVP.Text = "" Then Exit Sub
If txtMargemAV.Text = "" Then Exit Sub
If txtMargemAP.Text = "" Then Exit Sub
If txtCusto.Text = "" Then Exit Sub

'AUTONUMERAÇÃO
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_precos;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

Dim varMargemVV As Double
Dim varMargemVP As Double
Dim varMargemAV As Double
Dim varMargemAP As Double

varMargemVV = Left$(txtMargemVV.Text, Len(txtMargemVV.Text) - 1)
varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)
varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)
varMargemAP = Left$(txtMargemAP.Text, Len(txtMargemAP.Text) - 1)

sSQL = "INSERT INTO produtos_precos (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, MARGEM_VV, VALOR_VV, MARGEM_VP, VALOR_VP, MARGEM_AV, VALOR_AV, MARGEM_AP, VALOR_AP, CUSTO) VALUES (" & _
   AutoNumeracao & ", " & txtCodProduto.Text & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), " & txtCodProduto.Text & ", 'AJUSTE', " & Replace(CDbl(varMargemVV), ",", ".") & ", " & Replace(CCur(txtValorVV.Text), ",", ".") & ", " & Replace(CDbl(varMargemVP), ",", ".") & ", " & Replace(CCur(txtValorVP.Text), ",", ".") & ", " & Replace(CDbl(varMargemAV), ",", ".") & ", " & Replace(CCur(txtValorAV.Text), ",", ".") & ", " & Replace(CDbl(varMargemAP), ",", ".") & ", " & Replace(CCur(txtValorAP.Text), ",", ".") & ", " & Replace(CCur(txtCusto.Text), ",", ".") & "  );"
dbData.Execute sSQL
Debug.Print sSQL

LimparObjetos_Produtos
'cmdSalvar.Enabled = False
txtCodProduto.Text = ""
cmdSair_Click
End Sub

Private Sub CalcularPrecos()
Dim varValorCusto As Currency
If txtCusto.Text = "" Then Exit Sub
varValorCusto = txtCusto.Text

'CALCULAR PREÇO - VAREJO A VISTA
Dim varMargemVV As Currency
Dim varValorVV As Currency

If txtMargemVV.Text = "" Then Exit Sub

varMargemVV = Left$(txtMargemVV.Text, Len(txtMargemVV.Text) - 1)

varValorVV = (varValorCusto * varMargemVV) / 100
varValorVV = varValorCusto + varValorVV
txtValorVV.Text = Format(varValorVV, ocMONEY)

'CALCULAR PREÇO - VAREJO A PRAZO
Dim varMargemVP As Currency
Dim varValorVP As Currency

If txtMargemVP.Text = "" Then Exit Sub

varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)

varValorVP = (varValorCusto * varMargemVP) / 100
varValorVP = varValorCusto + varValorVP
txtValorVP.Text = Format(varValorVP, ocMONEY)

'CALCULAR PREÇO - ATACADO A VISTA
Dim varMargemAV As Currency
Dim varValorAV As Currency

If txtMargemAV.Text = "" Then Exit Sub

varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)

varValorAV = (varValorCusto * varMargemAV) / 100
varValorAV = varValorCusto + varValorAV
txtValorAV.Text = Format(varValorAV, ocMONEY)

'CALCULAR PREÇO - ATACADO A PRAZO
Dim varMargemAP As Currency
Dim varValorAP As Currency

If txtMargemAP.Text = "" Then Exit Sub

varMargemAP = Left$(txtMargemAP.Text, Len(txtMargemAP.Text) - 1)

varValorAP = (varValorCusto * varMargemAP) / 100
varValorAP = varValorCusto + varValorAP
txtValorAP.Text = Format(varValorAP, ocMONEY)
End Sub

Private Sub Form_Activate()
txtCusto.SetFocus
End Sub


Private Sub Form_Load()
cmdSalvar.Enabled = True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Produtos_AjustoPreco.Hide
Produtos_Estoque_Simples.Show 1
End Sub


Private Sub txtMargemAP_GotFocus()
If txtMargemAP.Text = "" Then Exit Sub
Dim varMargemAP As Currency

If Right(txtMargemAP.Text, 1) = "%" Then
   varMargemAP = Left$(txtMargemAP.Text, Len(txtMargemAP.Text) - 1)
Else
    varMargemAP = txtMargemAP.Text
End If

txtMargemAP.Text = varMargemAP

txtMargemAP.SelStart = 0
txtMargemAP.SelLength = Len(txtMargemAP.Text)
lblAviso.Visible = True
End Sub


Private Sub txtMargemAP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    If txtCusto.Text = "" Then Exit Sub
    varValorEstimado = Empty
    varCustoEstimado = CCur(txtCusto)
    Produtos_ValorEstimado.Show vbModal
    Unload Produtos_ValorEstimado
    txtMargemAP.Text = varValorEstimado
End If
End Sub


Private Sub txtMargemAP_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtMargemAP_LostFocus()
Dim varMargemAP As Currency

If txtMargemAP.Text = "" Then txtMargemAP.Text = 0
varMargemAP = txtMargemAP.Text

txtMargemAP.Text = FormatNumber(varMargemAP, 4) & "%"

CalcularPrecos
lblAviso.Visible = False
End Sub

Private Sub txtMargemAV_GotFocus()
If txtMargemAV.Text = "" Then Exit Sub
Dim varMargemAV As Currency

If Right(txtMargemAV.Text, 1) = "%" Then
   varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)
Else
    varMargemAV = txtMargemAV.Text
End If

txtMargemAV.Text = varMargemAV

txtMargemAV.SelStart = 0
txtMargemAV.SelLength = Len(txtMargemAV.Text)
lblAviso.Visible = True
End Sub


Private Sub txtMargemAV_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    If txtCusto.Text = "" Then Exit Sub
    varValorEstimado = Empty
    varCustoEstimado = CCur(txtCusto)
    Produtos_ValorEstimado.Show vbModal
    Unload Produtos_ValorEstimado
    txtMargemAV.Text = varValorEstimado
End If
End Sub


Private Sub txtMargemAV_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtMargemAV_LostFocus()
Dim varMargemAV As Currency

If txtMargemAV.Text = "" Then txtMargemAV.Text = 0
varMargemAV = txtMargemAV.Text

txtMargemAV.Text = FormatNumber(varMargemAV, 4) & "%"

CalcularPrecos
lblAviso.Visible = False
End Sub

Private Sub txtMargemVP_GotFocus()
If txtMargemVP.Text = "" Then Exit Sub
Dim varMargemVP As Currency

If Right(txtMargemVP.Text, 1) = "%" Then
   varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)
Else
    varMargemVP = txtMargemVP.Text
End If

txtMargemVP.Text = varMargemVP

txtMargemVP.SelStart = 0
txtMargemVP.SelLength = Len(txtMargemVP.Text)
lblAviso.Visible = True
End Sub


Private Sub txtMargemVP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    If txtCusto.Text = "" Then Exit Sub
    varValorEstimado = Empty
    varCustoEstimado = CCur(txtCusto)
    Produtos_ValorEstimado.Show vbModal
    Unload Produtos_ValorEstimado
    txtMargemVP.Text = varValorEstimado
End If
End Sub


Private Sub txtMargemVP_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtMargemVP_LostFocus()
Dim varMargemVP As Currency

If txtMargemVP.Text = "" Then txtMargemVP.Text = 0
varMargemVP = txtMargemVP.Text

txtMargemVP.Text = FormatNumber(varMargemVP, 4) & "%"

CalcularPrecos
lblAviso.Visible = False
End Sub

Private Sub txtMargemVV_GotFocus()
If txtMargemVV.Text = "" Then Exit Sub
Dim varMargemVV As Currency

If Right(txtMargemVV.Text, 1) = "%" Then
   varMargemVV = Left$(txtMargemVV.Text, Len(txtMargemVV.Text) - 1)
Else
    varMargemVV = txtMargemVV.Text
End If

txtMargemVV.Text = varMargemVV

txtMargemVV.SelStart = 0
txtMargemVV.SelLength = Len(txtMargemVV.Text)
lblAviso.Visible = True
End Sub

Private Sub txtMargemVV_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    If txtCusto.Text = "" Then Exit Sub
    varValorEstimado = Empty
    varCustoEstimado = CCur(txtCusto)
    Produtos_ValorEstimado.Show vbModal
    Unload Produtos_ValorEstimado
    txtMargemVV.Text = varValorEstimado
End If
End Sub


Private Sub txtMargemVV_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtMargemVV_LostFocus()
Dim varMargemVV As Currency

If txtMargemVV.Text = "" Then txtMargemVV.Text = 0
varMargemVV = txtMargemVV.Text

txtMargemVV.Text = FormatNumber(varMargemVV, 4) & "%"
If txtMargemVP.Text = "" Then txtMargemVP.Text = txtMargemVV.Text
If txtMargemAV.Text = "" Then txtMargemAV.Text = txtMargemVV.Text
If txtMargemAP.Text = "" Then txtMargemAP.Text = txtMargemVV.Text
CalcularPrecos
lblAviso.Visible = False
End Sub

Private Sub txtCusto_GotFocus()
txtCusto.SelStart = 0
txtCusto.SelLength = Len(txtCusto)
End Sub


Private Sub txtCusto_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtCusto_LostFocus()
Dim varLucro As Currency

If txtCusto.Text = "" Then Exit Sub
varLucro = txtCusto.Text

txtCusto.Text = FormatNumber(varLucro, 2)

CalcularPrecos
End Sub
Private Sub txtCodProduto_Change()
Dim sSQL As String
Dim r As ADODB.Recordset
If txtCodProduto.Text = "" Then Exit Sub

sSQL = "SELECT * FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

LimparObjetos_Produtos
MostrarDados_Produto r
MostrarObjetosPrecos
End Sub


Private Sub txtValorAP_Click()
SelectControl txtValorAP
End Sub

Private Sub txtValorAP_GotFocus()
SelectControl txtValorAP
End Sub

Private Sub txtValorAP_LostFocus()
If txtCusto.Text = "" Then Exit Sub
If txtValorAP.Text = "" Then Exit Sub

Dim a As Currency
Dim B As Currency
Dim c As Currency

a = txtCusto.Text
B = txtValorAP.Text
c = ((B - a) / a) * 100

txtMargemAP.Text = FormatNumber(c, 4) & "%"
End Sub


Private Sub txtValorAV_Click()
SelectControl txtValorAV
End Sub

Private Sub txtValorAV_GotFocus()
SelectControl txtValorAV
End Sub

Private Sub txtValorAV_LostFocus()
If txtCusto.Text = "" Then Exit Sub
If txtValorAV.Text = "" Then Exit Sub

Dim a As Currency
Dim B As Currency
Dim c As Currency

a = txtCusto.Text
B = txtValorAV.Text
c = ((B - a) / a) * 100

txtMargemAV.Text = FormatNumber(c, 4) & "%"
End Sub


Private Sub txtValorVP_Click()
SelectControl txtValorVP
End Sub

Private Sub txtValorVP_GotFocus()
SelectControl txtValorVP
End Sub

Private Sub txtValorVP_LostFocus()
If txtCusto.Text = "" Then Exit Sub
If txtValorVP.Text = "" Then Exit Sub

Dim a As Currency
Dim B As Currency
Dim c As Currency

a = txtCusto.Text
B = txtValorVP.Text
c = ((B - a) / a) * 100

txtMargemVP.Text = FormatNumber(c, 4) & "%"
End Sub


Private Sub txtValorVV_Click()
SelectControl txtValorVV
End Sub

Private Sub txtValorVV_GotFocus()
SelectControl txtValorVV
End Sub

Private Sub txtValorVV_LostFocus()
If txtCusto.Text = "" Then Exit Sub
If txtValorVV.Text = "" Then Exit Sub

Dim a As Currency
Dim B As Currency
Dim c As Currency

a = txtCusto.Text
B = txtValorVV.Text
c = ((B - a) / a) * 100

txtMargemVV.Text = FormatNumber(c, 4) & "%"
End Sub


