VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Aluguel_Cadastro_Equipamentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CADASTRO DE EQUIPAMENTOS"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   Icon            =   "Aluguel_Cadastro_Equipamentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4755
      Left            =   60
      TabIndex        =   10
      Top             =   1080
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   8387
      _Version        =   393216
      Tab             =   2
      TabHeight       =   450
      TabMaxWidth     =   2646
      WordWrap        =   0   'False
      TabCaption(0)   =   "Cadastro"
      TabPicture(0)   =   "Aluguel_Cadastro_Equipamentos.frx":23D2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdNovo"
      Tab(0).Control(1)=   "cmdSalvar"
      Tab(0).Control(2)=   "cmdExcluir"
      Tab(0).Control(3)=   "cmdAlterar"
      Tab(0).Control(4)=   "cmdCancelar"
      Tab(0).Control(5)=   "frmCadastro"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "Aluguel_Cadastro_Equipamentos.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Grid"
      Tab(1).Control(2)=   "cmdExibirContrato"
      Tab(1).Control(3)=   "lblTotal"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Contratos"
      TabPicture(2)   =   "Aluguel_Cadastro_Equipamentos.frx":240A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblEquipContratos"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdAbrirContranto"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Grid_Contrato"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   10515
         Begin VB.ComboBox cboConsCriterio 
            Height          =   315
            ItemData        =   "Aluguel_Cadastro_Equipamentos.frx":2426
            Left            =   3720
            List            =   "Aluguel_Cadastro_Equipamentos.frx":2428
            TabIndex        =   24
            Top             =   420
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.ComboBox cboOrganizar 
            Height          =   315
            ItemData        =   "Aluguel_Cadastro_Equipamentos.frx":242A
            Left            =   1980
            List            =   "Aluguel_Cadastro_Equipamentos.frx":242C
            TabIndex        =   22
            Top             =   420
            Width           =   1695
         End
         Begin VB.ComboBox cboCriterio 
            Height          =   315
            ItemData        =   "Aluguel_Cadastro_Equipamentos.frx":242E
            Left            =   60
            List            =   "Aluguel_Cadastro_Equipamentos.frx":2430
            TabIndex        =   20
            Top             =   420
            Width           =   1875
         End
         Begin ChamaleonBtn.chameleonButton cmdExibir 
            Height          =   555
            Left            =   7620
            TabIndex        =   26
            Top             =   300
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "Aluguel_Cadastro_Equipamentos.frx":2432
            PICN            =   "Aluguel_Cadastro_Equipamentos.frx":244E
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
            Height          =   555
            Left            =   9000
            TabIndex        =   27
            Top             =   300
            Width           =   1395
            _ExtentX        =   2461
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
            MICON           =   "Aluguel_Cadastro_Equipamentos.frx":41E0
            PICN            =   "Aluguel_Cadastro_Equipamentos.frx":41FC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblTituloCriterio 
            AutoSize        =   -1  'True
            Caption         =   "Organizar por:"
            Height          =   195
            Left            =   3720
            TabIndex        =   25
            Top             =   180
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Organizar por:"
            Height          =   195
            Left            =   1980
            TabIndex        =   23
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Critério:"
            Height          =   195
            Left            =   60
            TabIndex        =   21
            Top             =   180
            Width           =   525
         End
      End
      Begin VB.Frame frmCadastro 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   8295
         Begin VB.TextBox txtQuant 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   1080
            Width           =   1275
         End
         Begin VB.TextBox txtObs 
            Height          =   1455
            Left            =   120
            TabIndex        =   33
            Top             =   1740
            Width           =   7995
         End
         Begin VB.ComboBox cboFabricante 
            Height          =   315
            ItemData        =   "Aluguel_Cadastro_Equipamentos.frx":5F8E
            Left            =   5760
            List            =   "Aluguel_Cadastro_Equipamentos.frx":5F90
            TabIndex        =   2
            Top             =   480
            Width           =   2355
         End
         Begin VB.TextBox txtEquipamento 
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   3675
         End
         Begin VB.TextBox txtModelo 
            Height          =   315
            Left            =   3840
            TabIndex        =   1
            Top             =   480
            Width           =   1875
         End
         Begin VB.TextBox txtvalorDia 
            Height          =   315
            Left            =   1440
            TabIndex        =   4
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtValorHora 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2820
            TabIndex        =   5
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Observação"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   1500
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Valor Dia"
            Height          =   195
            Left            =   1440
            TabIndex        =   17
            Top             =   840
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   5760
            TabIndex        =   16
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Equipamento"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Modelo"
            Height          =   195
            Left            =   3840
            TabIndex        =   14
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Hora"
            Height          =   195
            Left            =   2820
            TabIndex        =   13
            Top             =   840
            Width           =   750
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   11
         Top             =   1380
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   5106
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   -66540
         TabIndex        =   28
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Cancelar"
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
         MICON           =   "Aluguel_Cadastro_Equipamentos.frx":5F92
         PICN            =   "Aluguel_Cadastro_Equipamentos.frx":5FAE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterar 
         Height          =   615
         Left            =   -66540
         TabIndex        =   29
         Top             =   2460
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Aluguel_Cadastro_Equipamentos.frx":7D40
         PICN            =   "Aluguel_Cadastro_Equipamentos.frx":7D5C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExcluir 
         Height          =   615
         Left            =   -66540
         TabIndex        =   30
         Top             =   3120
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Aluguel_Cadastro_Equipamentos.frx":9AEE
         PICN            =   "Aluguel_Cadastro_Equipamentos.frx":9B0A
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
         Left            =   -66540
         TabIndex        =   31
         Top             =   1140
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Salvar"
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
         MICON           =   "Aluguel_Cadastro_Equipamentos.frx":B89C
         PICN            =   "Aluguel_Cadastro_Equipamentos.frx":B8B8
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
         Left            =   -66540
         TabIndex        =   32
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Aluguel_Cadastro_Equipamentos.frx":D64A
         PICN            =   "Aluguel_Cadastro_Equipamentos.frx":D666
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Contrato 
         Height          =   3975
         Left            =   60
         TabIndex        =   36
         Top             =   660
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7011
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdExibirContrato 
         Height          =   255
         Left            =   -74880
         TabIndex        =   39
         Top             =   4380
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "&Exibir Contrato"
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
         MICON           =   "Aluguel_Cadastro_Equipamentos.frx":F3F8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAbrirContranto 
         Height          =   255
         Left            =   9300
         TabIndex        =   40
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "&Exibir Contrato"
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
         MICON           =   "Aluguel_Cadastro_Equipamentos.frx":F414
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblEquipContratos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EQUPAMENTOS"
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
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EQUPAMENTOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   11280
         TabIndex        =   37
         Top             =   4860
         Width           =   1890
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   -64620
         TabIndex        =   18
         Top             =   4440
         Width           =   225
      End
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10020
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   10725
      TabIndex        =   7
      Top             =   60
      Width           =   10755
      Begin VB.Image Image1 
         Height          =   480
         Left            =   180
         Picture         =   "Aluguel_Cadastro_Equipamentos.frx":F430
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EQUPAMENTOS"
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
         Left            =   900
         TabIndex        =   8
         Top             =   240
         Width           =   2430
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   9
      Top             =   5910
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14843
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "18:08"
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
End
Attribute VB_Name = "Aluguel_Cadastro_Equipamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Private printSQL As String
Dim sSQL As String
Dim r As ADODB.Recordset
Dim i As Integer
Private Function Inserir_Dados() As Boolean
Dim sSQL As String

sSQL = "INSERT INTO aluguel_cadastro_equipamento (COD_EQUIP, DESCRICAO, FABRICANTE, MODELO, VALOR_DIA, VALOR_HORA, ATIVO, ALUGADO, OBS, QUANT_ESTOQUE, QUANT_ALUGADA, CONTRATO) VALUES (" & _
   txtCodigo.Text & ", '" & txtEquipamento.Text & "', '" & cboFabricante.Text & "', '" & txtModelo.Text & "', " & Replace(CCur(txtvalorDia.Text), ",", ".") & ", " & Replace(CCur(txtValorHora.Text), ",", ".") & ", 1, 0, '" & txtObs.Text & "', " & txtQuant.Text & ", 0, 0);"

Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
Dim sSQL As String

sSQL = "UPDATE aluguel_cadastro_equipamento SET DESCRICAO = '" & txtEquipamento.Text & "', FABRICANTE = '" & cboFabricante.Text & "', MODELO = '" & txtModelo.Text & "', OBS = '" & txtObs.Text & "', VALOR_DIA = " & Replace(CCur(txtvalorDia.Text), ",", ".") & ", VALOR_HORA = " & Replace(CCur(txtValorHora.Text), ",", ".") & ", QUANT_ESTOQUE = " & txtQuant.Text & " WHERE (COD_EQUIP = " & txtCodigo.Text & ");"

Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub Auto_Numeracao()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT ISNULL(MAX(COD_EQUIP), 0) AS codigo FROM aluguel_cadastro_equipamento;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then txtCodigo.Text = r("codigo") + 1
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub FormatarGridContrato(rTabela As ADODB.Recordset)
Dim i As Integer, x As Integer
Dim j As Integer

With Grid_Contrato
   .Clear
   .Cols = 11
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 650
   .ColWidth(2) = 3000
   .ColWidth(3) = 650
   .ColWidth(4) = 900
   .ColWidth(5) = 650
   .ColWidth(6) = 900
   .ColWidth(7) = 900
   .ColWidth(8) = 900
   .ColWidth(9) = 900
   .ColWidth(10) = 900
   
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "CÓD"
   .TextMatrix(0, 2) = "CLIENTE"
   .TextMatrix(0, 3) = "TIPO"
   .TextMatrix(0, 4) = "VALOR"
   .TextMatrix(0, 5) = "QTDE"
   .TextMatrix(0, 6) = "SUBTOTAL"
   .TextMatrix(0, 7) = "DESC"
   .TextMatrix(0, 8) = "TOTAL"
   .TextMatrix(0, 9) = "INICIO"
   .TextMatrix(0, 10) = "FINAL"
   .Redraw = False
   
   i = 1
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("COD_LOCACAO")
         .TextMatrix(.rows - 1, 2) = rTabela("nome")
         .TextMatrix(.rows - 1, 3) = rTabela("TIPO_LOCACAO")
        .TextMatrix(.rows - 1, 4) = Format(rTabela("VALOR_UND"), ocMONEY)
        .TextMatrix(.rows - 1, 5) = rTabela("QUANT_ALUGADA")
         .TextMatrix(.rows - 1, 6) = Format(rTabela("TOTAL_ALUGADA"), ocMONEY)
         .TextMatrix(.rows - 1, 7) = Format(rTabela("DESCONTO"), ocMONEY)
         .TextMatrix(.rows - 1, 8) = Format(rTabela("VALOR_FINAL"), ocMONEY)
         .TextMatrix(.rows - 1, 9) = Format(rTabela("DATA_INICIO"), ocDATA2)
         .TextMatrix(.rows - 1, 10) = Format(rTabela("DATA_FINAL"), ocDATA2)


         rTabela.MoveNext
         
         .rows = .rows + 1
         i = i + 1
      Loop
   End If
   
    'For i = 1 To .rows - 1
    '   For j = 0 To .Cols - 1
    '      .Col = j
    '      .Row = i
    '
    '      If .TextMatrix(i, 7) = "ALUGADO" And .TextMatrix(i, 8) = "NÃO" Then
    '         .CellForeColor = vbBlue
    '      ElseIf .TextMatrix(i, 7) = "ALUGADO" And .TextMatrix(i, 8) = "SIM" Then
    '         .CellForeColor = vbGreen
    '      Else
    '         .CellForeColor = vbBlack
    '      End If
    '   Next
    'Next
   
   .rows = .rows - 1
   .Redraw = True

'lblTotal.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
End With
End Sub
Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer, x As Integer
Dim j As Integer

With Grid
   .Clear
   .Cols = 10
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 2500
   .ColWidth(3) = 1200
   .ColWidth(4) = 1300
   .ColWidth(5) = 950
   .ColWidth(6) = 950
   .ColWidth(7) = 1100
   .ColWidth(8) = 1100
   .ColWidth(9) = 1100
   
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "EQUIPAMENTO"
   .TextMatrix(0, 3) = "MODELO"
   .TextMatrix(0, 4) = "FABRICANTE"
   .TextMatrix(0, 5) = "V.DIÁRIA"
   .TextMatrix(0, 6) = "V.HORA"
   .TextMatrix(0, 7) = "ESTOQUE"
   .TextMatrix(0, 8) = "ALUGADO"
   .TextMatrix(0, 9) = "DISP."
   .Redraw = False
   
   i = 1
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("COD_EQUIP")
         .TextMatrix(.rows - 1, 2) = rTabela("DESCRICAO")
         .TextMatrix(.rows - 1, 3) = rTabela("MODELO")
         .TextMatrix(.rows - 1, 4) = rTabela("FABRICANTE")
         .TextMatrix(.rows - 1, 5) = Format(rTabela("VALOR_DIA"), ocMONEY)
         .TextMatrix(.rows - 1, 6) = Format(rTabela("VALOR_HORA"), ocMONEY)
         .TextMatrix(.rows - 1, 7) = rTabela("QUANT_ESTOQUE")
         .TextMatrix(.rows - 1, 8) = rTabela("QUANT_ALUGADA")
         .TextMatrix(.rows - 1, 9) = rTabela("VARQUANTDISPONIVEL")

         rTabela.MoveNext
         
         .rows = .rows + 1
         i = i + 1
      Loop
   End If
   
    For i = 1 To .rows - 1
       For j = 0 To .Cols - 1
          .Col = j
          .Row = i
          
          If .TextMatrix(i, 7) = "ALUGADO" And .TextMatrix(i, 8) = "NÃO" Then
             .CellForeColor = vbBlue
          ElseIf .TextMatrix(i, 7) = "ALUGADO" And .TextMatrix(i, 8) = "SIM" Then
             .CellForeColor = vbGreen
          Else
             .CellForeColor = vbBlack
          End If
       Next
    Next
   
   .rows = .rows - 1
   .Redraw = True

'lblTotal.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
End With
End Sub


Private Sub Limpar_Objetos()
txtCodigo.Text = ""
txtvalorDia.Text = ""
cboFabricante.Text = ""
txtEquipamento.Text = ""
txtModelo.Text = ""
txtValorHora.Text = ""
txtObs.Text = ""
txtQuant.Text = ""
End Sub


Private Sub cboConsCriterio_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
cboConsCriterio.Clear

If cboCriterio.Text = "EQUIPAMENTO" Then
    sSQL = "SELECT DISTINCT DESCRICAO FROM aluguel_cadastro_equipamento ORDER BY DESCRICAO;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       cboConsCriterio.AddItem r("DESCRICAO")
       r.MoveNext
    Loop
ElseIf cboCriterio.Text = "FABRICANTE" Then
    sSQL = "SELECT DISTINCT FABRICANTE FROM aluguel_cadastro_equipamento ORDER BY FABRICANTE;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       cboConsCriterio.AddItem r("FABRICANTE")
       r.MoveNext
    Loop
ElseIf cboCriterio.Text = "MODELO" Then
    sSQL = "SELECT DISTINCT MODELO FROM aluguel_cadastro_equipamento ORDER BY MODELO;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       cboConsCriterio.AddItem r("MODELO")
       r.MoveNext
    Loop
End If

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboConsCriterio
End Sub


Private Sub cboCriterio_Click()
cboCriterio_LostFocus
End Sub

Private Sub cboFabricante_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

'Limpa a lista atual
cboFabricante.Clear

sSQL = "SELECT DISTINCT fabricante FROM aluguel_cadastro_equipamento ORDER BY fabricante;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboFabricante.AddItem ValidateNull(r("fabricante"))
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboFabricante
End Sub





Private Sub cboCriterio_GotFocus()
cboCriterio.Clear
cboCriterio.AddItem "TODOS"
cboCriterio.AddItem "EQUIPAMENTO"
cboCriterio.AddItem "FABRICANTE"
cboCriterio.AddItem "MODELO"
moCombo.AttachTo cboCriterio
End Sub


Private Sub cboCriterio_LostFocus()
If cboCriterio.Text = "TODOS" Then
    lblTituloCriterio.Visible = False
    cboConsCriterio.Visible = False
    cboConsCriterio.Text = ""
ElseIf cboCriterio.Text = "EQUIPAMENTO" Then
    cboConsCriterio.Text = ""
    lblTituloCriterio.Visible = True
    lblTituloCriterio.Caption = "Equipamento"
    cboConsCriterio.Visible = True
ElseIf cboCriterio.Text = "FABRICANTE" Then
    cboConsCriterio.Text = ""
    lblTituloCriterio.Visible = True
    lblTituloCriterio.Caption = "Fabricante"
    cboConsCriterio.Visible = True
ElseIf cboCriterio.Text = "MODELO" Then
    cboConsCriterio.Text = ""
    lblTituloCriterio.Visible = True
    lblTituloCriterio.Caption = "Modelo"
    cboConsCriterio.Visible = True
Else
    lblTituloCriterio.Visible = False
    cboConsCriterio.Visible = False
    cboConsCriterio.Text = ""
End If
End Sub


Private Sub cboFabricante_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboOrganizar_GotFocus()
cboOrganizar.Clear
cboOrganizar.AddItem "EQUIPAMENTO"
cboOrganizar.AddItem "MODELO"
cboOrganizar.AddItem "FABRICANTE"
moCombo.AttachTo cboOrganizar
End Sub


Private Sub chameleonButton1_Click()

End Sub

Private Sub cmdAbrirContranto_Click()
Dim varCodContrato As Integer
i = Grid_Contrato.Row

'If txtCodCliente.Text = "" Then MsgBox "Escolha um cliente!", vbInformation, "Aviso do Sistema": Exit Sub

varCodContrato = Grid_Contrato.TextMatrix(i, 1)

'If ShowMsg("Deseja atualizar o cliente " & cboCliente.Text & " ?", vbInformation + vbYesNo) = vbYes Then
    Load Aluguel_Cadastro
    Aluguel_Cadastro.SSTab1.Tab = 0
    Aluguel_Cadastro.cmdNovo.Enabled = False
    Aluguel_Cadastro.cmdSalvar.Enabled = False
    Aluguel_Cadastro.cmdCancelar.Enabled = False
    Aluguel_Cadastro.lblCodigo.Caption = varCodContrato
    Aluguel_Cadastro.Show 1
'End If
End Sub


Private Sub cmdAlterar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Or txtEquipamento.Text = "" Then Exit Sub

If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Limpar_Objetos
Form_Load
End Sub

Private Sub cmdCancelar_Click()
Limpar_Objetos
Form_Load
End Sub

Private Sub cmdExcluir_Click()
Dim sSQL As String
Dim bRet As Boolean

If txtCodigo.Text = "" Then Exit Sub

If ShowMsg("Excluir esse equipamento?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

sSQL = "DELETE FROM aluguel_cadastro_equipamento WHERE (COD_EQUIP = " & txtCodigo.Text & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

Limpar_Objetos
Form_Load
End Sub

Private Sub cmdExibir_Click()
If cboCriterio.Text = "" Then Exit Sub

Dim INDICE As String
If cboOrganizar.Text = "EQUIPAMENTO" Then
   INDICE = "DESCRICAO "
ElseIf cboOrganizar.Text = "MODELO" Then
   INDICE = "MODELO "
ElseIf cboOrganizar.Text = "FABRICANTE" Then
   INDICE = "FABRICANTE "
Else
   INDICE = "DESCRICAO "
End If

'    sSQL = "SELECT *, (CASE WHEN ALUGADO = 1 THEN 'ALUGADO' ELSE 'LIVRE' END) AS varSituacao,(CASE WHEN RESERVADO = 1 THEN 'SIM' ELSE 'NÃO' END) as varReservado FROM aluguel_cadastro_equipamento ORDER BY  " & INDICE

If cboCriterio.Text = "TODOS" Then
    sSQL = "SELECT COD_EQUIP, DESCRICAO, FABRICANTE, MODELO, VALOR_DIA, VALOR_HORA, QUANT_ESTOQUE, QUANT_ALUGADA, (QUANT_ESTOQUE - QUANT_ALUGADA) as varQuantDisponivel FROM aluguel_cadastro_equipamento ORDER BY  " & INDICE
ElseIf cboCriterio.Text = "EQUIPAMENTO" Then
    If cboConsCriterio.Text = "" Then Exit Sub
    sSQL = "SELECT COD_EQUIP, DESCRICAO, FABRICANTE, MODELO, VALOR_DIA, VALOR_HORA, QUANT_ESTOQUE, QUANT_ALUGADA, (QUANT_ESTOQUE - QUANT_ALUGADA) as varQuantDisponivel FROM aluguel_cadastro_equipamento where (DESCRICAO = '" & cboConsCriterio.Text & "') ORDER BY  " & INDICE
ElseIf cboCriterio.Text = "MODELO" Then
    If cboConsCriterio.Text = "" Then Exit Sub
    sSQL = "SELECT COD_EQUIP, DESCRICAO, FABRICANTE, MODELO, VALOR_DIA, VALOR_HORA, QUANT_ESTOQUE, QUANT_ALUGADA, (QUANT_ESTOQUE - QUANT_ALUGADA) as varQuantDisponivel FROM aluguel_cadastro_equipamento where (MODELO = '" & cboConsCriterio.Text & "') ORDER BY  " & INDICE
ElseIf cboCriterio.Text = "FABRICANTE" Then
    If cboConsCriterio.Text = "" Then Exit Sub
    sSQL = "SELECT COD_EQUIP, DESCRICAO, FABRICANTE, MODELO, VALOR_DIA, VALOR_HORA, QUANT_ESTOQUE, QUANT_ALUGADA, (QUANT_ESTOQUE - QUANT_ALUGADA) as varQuantDisponivel FROM aluguel_cadastro_equipamento where (FABRICANTE = '" & cboConsCriterio.Text & "') ORDER BY  " & INDICE
End If

Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
printSQL = sSQL

End Sub
Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To var_Grid.rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub cmdExibirContrato_Click()
'If cboCriterio.Text = "" Then Exit Sub

'Dim INDICE As String
'If cboOrganizar.Text = "EQUIPAMENTO" Then
'   INDICE = "DESCRICAO "
'ElseIf cboOrganizar.Text = "MODELO" Then
'   INDICE = "MODELO "
'ElseIf cboOrganizar.Text = "FABRICANTE" Then
'   INDICE = "FABRICANTE "
'Else
'   INDICE = "DESCRICAO "
'End If

i = Grid.Row

'If GridProdutos.TextMatrix(i, 1) = "SIM" Then
'    MsgBox "Esse equipamento já foi devolvido!", vbInformation, "Aviso do Sistema"
'    Exit Sub
'End If


'If cboCriterio.Text = "TODOS" Then
    'sSQL = "SELECT COD_EQUIP, DESCRICAO, FABRICANTE, MODELO, VALOR_DIA, VALOR_HORA, QUANT_ESTOQUE, QUANT_ALUGADA, (QUANT_ESTOQUE - QUANT_ALUGADA) as varQuantDisponivel FROM aluguel_cadastro_equipamento ORDER BY  " & INDICE
    SSTab1.Tab = 2
    
    lblEquipContratos.Caption = Grid.TextMatrix(i, 2)

    sSQL = "SELECT Aluguel_Cadastro_Itens.CODIGO, Aluguel_Cadastro_Itens.COD_LOCACAO, Aluguel_Cadastro_Itens.COD_EQUIP, Aluguel_Cadastro_Itens.TIPO_LOCACAO, " & _
                "Aluguel_Cadastro_Itens.DATA_INICIO, Aluguel_Cadastro_Itens.DATA_FINAL, Aluguel_Cadastro_Itens.VALOR_UND, Aluguel_Cadastro_Itens.QUANT_ALUGADA, " & _
                "Aluguel_Cadastro_Itens.TOTAL_ALUGADA , Aluguel_Cadastro_Itens.DESCONTO, Aluguel_Cadastro_Itens.VALOR_FINAL, Cliente.nome " & _
            "From Aluguel_Cadastro_Itens INNER JOIN " & _
                "Aluguel_Cadastro ON Aluguel_Cadastro_Itens.COD_LOCACAO = Aluguel_Cadastro.CODIGO INNER JOIN " & _
                "cliente ON Aluguel_Cadastro.COD_CLIENTE = cliente.CODIGO " & _
            "Where (Aluguel_Cadastro_Itens.COD_EQUIP = " & Grid.TextMatrix(i, 1) & ") And (Aluguel_Cadastro_Itens.EXCLUIDO = 0) And (Aluguel_Cadastro_Itens.DEVOLVIDO = 0)"

'ElseIf cboCriterio.Text = "EQUIPAMENTO" Then
'    If cboConsCriterio.Text = "" Then Exit Sub
'    sSQL = "SELECT COD_EQUIP, DESCRICAO, FABRICANTE, MODELO, VALOR_DIA, VALOR_HORA, QUANT_ESTOQUE, QUANT_ALUGADA, (QUANT_ESTOQUE - QUANT_ALUGADA) as varQuantDisponivel FROM aluguel_cadastro_equipamento where (DESCRICAO = '" & cboConsCriterio.Text & "') ORDER BY  " & INDICE
'ElseIf cboCriterio.Text = "MODELO" Then
'    If cboConsCriterio.Text = "" Then Exit Sub
'    sSQL = "SELECT COD_EQUIP, DESCRICAO, FABRICANTE, MODELO, VALOR_DIA, VALOR_HORA, QUANT_ESTOQUE, QUANT_ALUGADA, (QUANT_ESTOQUE - QUANT_ALUGADA) as varQuantDisponivel FROM aluguel_cadastro_equipamento where (MODELO = '" & cboConsCriterio.Text & "') ORDER BY  " & INDICE
'ElseIf cboCriterio.Text = "FABRICANTE" Then
'    If cboConsCriterio.Text = "" Then Exit Sub
'    sSQL = "SELECT COD_EQUIP, DESCRICAO, FABRICANTE, MODELO, VALOR_DIA, VALOR_HORA, QUANT_ESTOQUE, QUANT_ALUGADA, (QUANT_ESTOQUE - QUANT_ALUGADA) as varQuantDisponivel FROM aluguel_cadastro_equipamento where (FABRICANTE = '" & cboConsCriterio.Text & "') ORDER BY  " & INDICE
'End If

Set r = dbData.OpenRecordset(sSQL)

FormatarGridContrato r

If r.State <> 0 Then r.Close
Set r = Nothing
printSQL = sSQL
End Sub

Private Sub cmdImprimir_Click()
'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini
Dim r As ADODB.Recordset

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Aluguel_Listadeequip.Relatorio.Recordset = r
'REL_Aluguel_Listadeequip.dfQuant.Caption = "QUANTIDADE: " & txtCONquant.Text
REL_Aluguel_Listadeequip.dfData.Caption = Format(Date, "dd/mm/yy")
REL_Aluguel_Listadeequip.lblTitulo.Caption = "RELATÓRIO DE EQUIPAMENTOS"

'If cboFiltro.Text = "TODOS" Then
'   REL_Aluguel_Listadeequip.dfTipo.Caption = "Tipo: Todos os registros"
'ElseIf cboFiltro.Text = "PERIODO" Then
'   REL_Aluguel_Listadeequip.dfTipo.Caption = "Tipo: Intervalo de " & Mask1.Text & " à " & Mask2.Text
'ElseIf cboFiltro.Text = "MÊS" Then
'   REL_Aluguel_Listadeequip.dfTipo.Caption = "Tipo: Mês = " & cboMes.Text & "/" & cboAno.Text
'ElseIf cboFiltro.Text = "CLIENTE" Then
'   REL_Aluguel_Listadeequip.dfTipo.Caption = "Cliente = " & cboNome.Text
'Else
'   REL_Aluguel_Listadeequip.dfTipo.Caption = "Tipo:"
'End If

REL_Aluguel_Listadeequip.Relatorio.NomeImpressora = var_Impressora
REL_Aluguel_Listadeequip.Relatorio.Ativar
Unload REL_Aluguel_Listadeequip

Me.Show 1

End Sub

Private Sub cmdNovo_Click()
Limpar_Objetos
Auto_Numeracao
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = False
frmCadastro.Enabled = True
txtEquipamento.SetFocus
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo TrataErro

If txtCodigo.Text = "" Or txtEquipamento.Text = "" Then Exit Sub

If Not Inserir_Dados Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Limpar_Objetos
Form_Load
Exit Sub
   
TrataErro:
   If Err.Number = 3022 Then
      ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
MostrarGrid
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdNovo.Enabled = True
frmCadastro.Enabled = False
SSTab1.Tab = 0
End Sub

Private Sub Mostrar_equipamento(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then
   cboFabricante.Text = rTabela("FABRICANTE")
   txtModelo.Text = rTabela("MODELO")
   txtObs.Text = ValidateNull(rTabela("OBS"))
   txtEquipamento.Text = rTabela("DESCRICAO")
   txtQuant.Text = rTabela("QUANT_ESTOQUE")
   txtvalorDia.Text = Format(rTabela("VALOR_DIA"), ocMONEY)
   txtValorHora.Text = Format(rTabela("VALOR_HORA"), ocMONEY)
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
frmCadastro.Enabled = True
txtCodigo.Text = ""
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub

Private Sub MostrarGrid()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT COD_EQUIP, DESCRICAO, FABRICANTE, MODELO, VALOR_DIA, VALOR_HORA, QUANT_ESTOQUE, QUANT_ALUGADA, (QUANT_ESTOQUE - QUANT_ALUGADA) as varQuantDisponivel FROM aluguel_cadastro_equipamento ORDER BY descricao;"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r
printSQL = sSQL

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub txtCodigo_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

If cmdAlterar.Enabled = True Then
   sSQL = "SELECT * FROM aluguel_cadastro_equipamento WHERE (COD_EQUIP = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then Mostrar_equipamento r
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If

SSTab1.Tab = 0
End Sub

Private Sub txtEquipamento_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtModelo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtvalorDia_GotFocus()
SelectControl txtvalorDia
End Sub

Private Sub txtvalorDia_LostFocus()
If txtvalorDia.Text = "" Then
   txtvalorDia.Text = Format(0, ocMONEY)
Else
   txtvalorDia.Text = Format(txtvalorDia, ocMONEY)
End If
End Sub


Private Sub txtValorHora_GotFocus()
SelectControl txtValorHora
End Sub

Private Sub txtValorHora_LostFocus()
If txtValorHora.Text = "" Then
   txtValorHora.Text = Format(0, ocMONEY)
Else
   txtValorHora.Text = Format(txtValorHora, ocMONEY)
End If
End Sub


