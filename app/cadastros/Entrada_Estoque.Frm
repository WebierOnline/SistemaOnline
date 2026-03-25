VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Entrada_Estoque 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTRADA NO ESTOQUE PELA XML"
   ClientHeight    =   9930
   ClientLeft      =   735
   ClientTop       =   1455
   ClientWidth     =   16005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Frm_NF"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9930
   ScaleWidth      =   16005
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Frm_NF 
      Height          =   9495
      Left            =   60
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   120
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Nota Fiscal"
      TabPicture(0)   =   "Entrada_Estoque.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdBuscarXML"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdExcluir"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ChaDownloadXML"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdImportarXML"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSair"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdNovo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmNota"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "frmValores"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCodNota"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCodPedido"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Tab_Entrada"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "frmProdFracionado"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "Entrada_Estoque.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label19"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label20"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label21"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label22"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label18"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label17"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label10"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label12"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label13"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label9"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblTotalNota"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lblSomaNota"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "GridNotas"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Frame3"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      Begin VB.Frame frmProdFracionado 
         Height          =   3675
         Left            =   3420
         TabIndex        =   219
         Top             =   720
         Visible         =   0   'False
         Width           =   9615
         Begin ChamaleonBtn.chameleonButton cmdProdFrancionado 
            Height          =   615
            Left            =   1740
            TabIndex        =   225
            Top             =   2820
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1085
            BTYPE           =   3
            TX              =   "Francionado"
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
            MICON           =   "Entrada_Estoque.frx":0038
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdProdOriginal 
            Height          =   615
            Left            =   5040
            TabIndex        =   226
            Top             =   2820
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1085
            BTYPE           =   3
            TX              =   "Original"
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
            MICON           =   "Entrada_Estoque.frx":0054
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label72 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Você deseja deseja fazer a entrada no estoque com o produto relacionado ou com o produto original? "
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
            Height          =   720
            Left            =   240
            TabIndex        =   224
            Top             =   1860
            Width           =   9225
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblNomeProdRel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   360
            Left            =   180
            TabIndex        =   223
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label70 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Está fracionado com o produto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   180
            TabIndex        =   222
            Top             =   900
            Width           =   5445
         End
         Begin VB.Label lblNomeProdSel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   360
            Left            =   180
            TabIndex        =   221
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label69 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Esse produto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   180
            TabIndex        =   220
            Top             =   120
            Width           =   2445
         End
      End
      Begin TabDlg.SSTab Tab_Entrada 
         Height          =   3195
         Left            =   120
         TabIndex        =   147
         Top             =   1380
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   5636
         _Version        =   393216
         TabHeight       =   520
         TabMaxWidth     =   4057
         TabCaption(0)   =   "Produtos"
         TabPicture(0)   =   "Entrada_Estoque.frx":0070
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblProdutoCadastrado"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frmItens"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtCFOPSel"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtCSTSel"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtNCMSel"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtDescricaoSel"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtValorSel"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtCodBarraSel"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtUnidMedSel"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtQuantSel"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Informações"
         TabPicture(1)   =   "Entrada_Estoque.frx":008C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).Control(1)=   "Frame13"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Parcelas e Transportes"
         TabPicture(2)   =   "Entrada_Estoque.frx":00A8
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame2"
         Tab(2).Control(1)=   "Frame18"
         Tab(2).ControlCount=   2
         Begin VB.Frame Frame18 
            Caption         =   "Parcelas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2715
            Left            =   -66540
            TabIndex        =   217
            Top             =   360
            Width           =   7035
            Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
               Height          =   1935
               Left            =   120
               TabIndex        =   218
               Top             =   420
               Width           =   6795
               _ExtentX        =   11986
               _ExtentY        =   3413
               _Version        =   393216
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Informações"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2715
            Left            =   -66600
            TabIndex        =   211
            Top             =   360
            Width           =   7095
            Begin VB.TextBox txtInformacao 
               Alignment       =   1  'Right Justify
               Height          =   855
               Left            =   120
               MaxLength       =   50
               TabIndex        =   215
               Top             =   540
               Width           =   6840
            End
            Begin VB.Label Label68 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Informações Adicionais"
               Height          =   195
               Left            =   120
               TabIndex        =   216
               Top             =   300
               Width           =   1635
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Transporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2715
            Left            =   -74940
            TabIndex        =   184
            Top             =   360
            Width           =   8355
            Begin VB.TextBox txtPlaca 
               Height          =   315
               Left            =   6480
               MaxLength       =   8
               TabIndex        =   204
               Top             =   1140
               Width           =   1245
            End
            Begin VB.TextBox txtPlacaUF 
               Height          =   315
               Left            =   7800
               MaxLength       =   2
               TabIndex        =   203
               Top             =   1140
               Width           =   465
            End
            Begin VB.Frame Frame12 
               Caption         =   "Volume"
               Height          =   1095
               Left            =   60
               TabIndex        =   190
               Top             =   1500
               Width           =   8235
               Begin VB.TextBox txtVolPesoLiquido 
                  Height          =   315
                  Left            =   6480
                  MaxLength       =   50
                  TabIndex        =   196
                  Top             =   480
                  Width           =   1545
               End
               Begin VB.TextBox txtVolNumeracao 
                  Height          =   315
                  Left            =   3840
                  MaxLength       =   50
                  TabIndex        =   195
                  Top             =   480
                  Width           =   1065
               End
               Begin VB.TextBox txtVolMarca 
                  Height          =   315
                  Left            =   2160
                  MaxLength       =   50
                  TabIndex        =   194
                  Top             =   480
                  Width           =   1665
               End
               Begin VB.TextBox txtVolEspecie 
                  Height          =   315
                  Left            =   900
                  MaxLength       =   50
                  TabIndex        =   193
                  Top             =   480
                  Width           =   1245
               End
               Begin VB.TextBox txtVolQuant 
                  Height          =   315
                  Left            =   60
                  MaxLength       =   50
                  TabIndex        =   192
                  Top             =   480
                  Width           =   825
               End
               Begin VB.TextBox txtVolPesoBruto 
                  Height          =   315
                  Left            =   4920
                  MaxLength       =   50
                  TabIndex        =   191
                  Top             =   480
                  Width           =   1545
               End
               Begin VB.Label Label42 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Numeração"
                  Height          =   195
                  Left            =   3840
                  TabIndex        =   202
                  Top             =   240
                  Width           =   825
               End
               Begin VB.Label Label40 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Marca"
                  Height          =   195
                  Left            =   2160
                  TabIndex        =   201
                  Top             =   240
                  Width           =   450
               End
               Begin VB.Label Label36 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Qtde Vol."
                  Height          =   195
                  Left            =   60
                  TabIndex        =   200
                  Top             =   240
                  Width           =   660
               End
               Begin VB.Label Label35 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Peso bruto"
                  Height          =   195
                  Left            =   4920
                  TabIndex        =   199
                  Top             =   240
                  Width           =   765
               End
               Begin VB.Label Label34 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Peso liquido"
                  Height          =   195
                  Left            =   6480
                  TabIndex        =   198
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label33 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Espécie"
                  Height          =   195
                  Left            =   900
                  TabIndex        =   197
                  Top             =   240
                  Width           =   570
               End
            End
            Begin VB.TextBox txtCodTransporte 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5280
               MaxLength       =   50
               TabIndex        =   188
               Top             =   840
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.ComboBox cboTransporte 
               Height          =   315
               Left            =   120
               TabIndex        =   187
               Top             =   1140
               Width           =   6315
            End
            Begin VB.ComboBox cboModFrete 
               Height          =   315
               Left            =   120
               TabIndex        =   185
               Top             =   540
               Width           =   3375
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Placa"
               Height          =   195
               Left            =   6480
               TabIndex        =   206
               Top             =   900
               Width           =   405
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "UF"
               Height          =   195
               Left            =   7800
               TabIndex        =   205
               Top             =   900
               Width           =   210
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Transportadora"
               Height          =   195
               Left            =   120
               TabIndex        =   189
               Top             =   900
               Width           =   1080
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Frete"
               Height          =   195
               Left            =   120
               TabIndex        =   186
               Top             =   300
               Width           =   945
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Valores e Impostos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2715
            Left            =   -74880
            TabIndex        =   163
            Top             =   360
            Width           =   8235
            Begin VB.TextBox txtValorCOFINS 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1695
               MaxLength       =   50
               TabIndex        =   209
               Top             =   1680
               Width           =   1560
            End
            Begin VB.TextBox txtValorPIS 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   60
               MaxLength       =   50
               TabIndex        =   207
               Top             =   1680
               Width           =   1560
            End
            Begin VB.TextBox txtValorFrete 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   60
               MaxLength       =   50
               TabIndex        =   173
               Top             =   1080
               Width           =   1560
            End
            Begin VB.TextBox txtBaseICMS 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   60
               MaxLength       =   50
               TabIndex        =   172
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtBaseICMSST 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3315
               MaxLength       =   50
               TabIndex        =   171
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtValorSeguro 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1695
               MaxLength       =   50
               TabIndex        =   170
               Top             =   1080
               Width           =   1560
            End
            Begin VB.TextBox txtValorOutrasDespesas 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4935
               MaxLength       =   50
               TabIndex        =   169
               Top             =   1080
               Width           =   1560
            End
            Begin VB.TextBox txtValorICMS 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1695
               MaxLength       =   50
               TabIndex        =   168
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtValorICMSST 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4935
               MaxLength       =   50
               TabIndex        =   167
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtValorIPI 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6555
               MaxLength       =   50
               TabIndex        =   166
               Top             =   1080
               Width           =   1560
            End
            Begin VB.TextBox txtValorDesconto 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3315
               MaxLength       =   50
               TabIndex        =   165
               Top             =   1080
               Width           =   1560
            End
            Begin VB.TextBox txtTotaldosProdutos 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6555
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   164
               TabStop         =   0   'False
               Top             =   480
               Width           =   1560
            End
            Begin VB.Label Label66 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor do COFINS"
               Height          =   195
               Left            =   1740
               TabIndex        =   210
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor do PIS"
               Height          =   195
               Left            =   60
               TabIndex        =   208
               Top             =   1440
               Width           =   885
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Base ICMS"
               Height          =   195
               Left            =   60
               TabIndex        =   183
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Base ICMS ST"
               Height          =   195
               Left            =   3315
               TabIndex        =   182
               Top             =   240
               Width           =   1050
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor do Seguro"
               Height          =   195
               Left            =   1695
               TabIndex        =   181
               Top             =   840
               Width           =   1140
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Outras Despesas"
               Height          =   195
               Left            =   4935
               TabIndex        =   180
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor do Frete"
               Height          =   195
               Left            =   60
               TabIndex        =   179
               Top             =   840
               Width           =   990
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor do ICMS"
               Height          =   195
               Left            =   1695
               TabIndex        =   178
               Top             =   240
               Width           =   1020
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor do ICMS ST"
               Height          =   195
               Left            =   4935
               TabIndex        =   177
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor do IPI"
               Height          =   195
               Left            =   6555
               TabIndex        =   176
               Top             =   840
               Width           =   825
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor do Desconto"
               Height          =   195
               Left            =   3315
               TabIndex        =   175
               Top             =   840
               Width           =   1320
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total dos Produtos"
               Height          =   195
               Left            =   6555
               TabIndex        =   174
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox txtQuantSel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   7560
            TabIndex        =   162
            ToolTipText     =   "Quant."
            Top             =   2760
            Width           =   825
         End
         Begin VB.TextBox txtUnidMedSel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   6960
            TabIndex        =   161
            ToolTipText     =   "Unid. Med."
            Top             =   2760
            Width           =   570
         End
         Begin VB.TextBox txtCodBarraSel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   120
            TabIndex        =   160
            ToolTipText     =   "EAN"
            Top             =   2760
            Width           =   1590
         End
         Begin VB.TextBox txtValorSel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   8400
            TabIndex        =   159
            ToolTipText     =   "Valor de Compra"
            Top             =   2760
            Width           =   1290
         End
         Begin VB.TextBox txtDescricaoSel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   1740
            TabIndex        =   158
            ToolTipText     =   "Descrição"
            Top             =   2760
            Width           =   5190
         End
         Begin VB.TextBox txtNCMSel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   9780
            TabIndex        =   157
            ToolTipText     =   "NCM"
            Top             =   2760
            Width           =   1170
         End
         Begin VB.TextBox txtCSTSel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   11760
            TabIndex        =   156
            ToolTipText     =   "CST"
            Top             =   2760
            Width           =   570
         End
         Begin VB.TextBox txtCFOPSel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   10980
            TabIndex        =   155
            ToolTipText     =   "CFOP"
            Top             =   2760
            Width           =   750
         End
         Begin VB.Frame frmItens 
            Caption         =   "Itens da Nota"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2355
            Left            =   120
            TabIndex        =   148
            Top             =   360
            Width           =   15375
            Begin MSFlexGridLib.MSFlexGrid GridNotasItens 
               Height          =   1755
               Left            =   120
               TabIndex        =   149
               Top             =   240
               Width           =   15135
               _ExtentX        =   26696
               _ExtentY        =   3096
               _Version        =   393216
               BackColor       =   16777215
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin ChamaleonBtn.chameleonButton cmdExcluirItem 
               Height          =   255
               Left            =   120
               TabIndex        =   150
               ToolTipText     =   "Adiciona"
               Top             =   2025
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   "Excluir"
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
               MICON           =   "Entrada_Estoque.frx":00C4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblValorNota 
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
               Left            =   14700
               TabIndex        =   154
               Top             =   2100
               Width           =   225
            End
            Begin VB.Label lblSubTotal 
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
               Left            =   12420
               TabIndex        =   153
               Top             =   2100
               Width           =   225
            End
            Begin VB.Label lblTotalDesc 
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
               Left            =   13800
               TabIndex        =   152
               Top             =   2100
               Width           =   225
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
               Height          =   195
               Left            =   9900
               TabIndex        =   151
               Top             =   2100
               Visible         =   0   'False
               Width           =   225
            End
         End
         Begin VB.Label lblProdutoCadastrado 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   12420
            TabIndex        =   227
            Top             =   2820
            Width           =   75
         End
      End
      Begin VB.TextBox txtCodPedido 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   13920
         TabIndex        =   113
         TabStop         =   0   'False
         ToolTipText     =   "Cód do Pedido"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtCodNota 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   14820
         TabIndex        =   112
         TabStop         =   0   'False
         ToolTipText     =   "Cód da Nota"
         Top             =   360
         Width           =   855
      End
      Begin VB.Frame frmValores 
         Caption         =   "Produto Selecionado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   120
         TabIndex        =   97
         Top             =   4560
         Width           =   13395
         Begin VB.TextBox txtCodProdExist 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   1320
            TabIndex        =   129
            ToolTipText     =   "EAN"
            Top             =   180
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.TextBox txtQuantExist 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   121
            ToolTipText     =   "Quant."
            Top             =   240
            Width           =   825
         End
         Begin VB.TextBox txtUnidMedExist 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   120
            ToolTipText     =   "Unid. Med."
            Top             =   240
            Width           =   570
         End
         Begin VB.TextBox txtCodBarraExist 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   119
            ToolTipText     =   "EAN"
            Top             =   240
            Width           =   1590
         End
         Begin VB.TextBox txtValorExist 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   118
            ToolTipText     =   "Valor de Venda"
            Top             =   240
            Width           =   1290
         End
         Begin VB.TextBox txtDescricaoExist 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   117
            ToolTipText     =   "Descrição"
            Top             =   240
            Width           =   5190
         End
         Begin VB.TextBox txtNCMExist 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   9780
            Locked          =   -1  'True
            TabIndex        =   116
            ToolTipText     =   "NCM"
            Top             =   240
            Width           =   1170
         End
         Begin VB.TextBox txtCSTExist 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   11760
            Locked          =   -1  'True
            TabIndex        =   115
            ToolTipText     =   "CST"
            Top             =   240
            Width           =   570
         End
         Begin VB.TextBox txtCFOPExist 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   10980
            Locked          =   -1  'True
            TabIndex        =   114
            ToolTipText     =   "CFOP"
            Top             =   240
            Width           =   750
         End
         Begin VB.Frame frmVenda 
            Caption         =   "Venda"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   5280
            TabIndex        =   104
            Top             =   1080
            Width           =   8055
            Begin VB.Frame Frame14 
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
               Left            =   60
               TabIndex        =   77
               Top             =   240
               Width           =   2115
               Begin VB.TextBox txtValorVV 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1020
                  TabIndex        =   15
                  Top             =   480
                  Width           =   795
               End
               Begin VB.TextBox txtMargemVV 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   60
                  TabIndex        =   14
                  Top             =   480
                  Width           =   895
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton1 
                  Height          =   315
                  Left            =   1860
                  TabIndex        =   16
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
                  MICON           =   "Entrada_Estoque.frx":00E0
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label54 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor"
                  Height          =   195
                  Left            =   1020
                  TabIndex        =   81
                  Top             =   240
                  Width           =   360
               End
               Begin VB.Label Label53 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lucro %"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   85
                  Top             =   240
                  Width           =   570
               End
            End
            Begin VB.Frame Frame15 
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
               Left            =   2220
               TabIndex        =   111
               Top             =   240
               Width           =   1875
               Begin VB.TextBox txtMargemVP 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   60
                  TabIndex        =   17
                  Top             =   480
                  Width           =   895
               End
               Begin VB.TextBox txtValorVP 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1020
                  TabIndex        =   18
                  Top             =   480
                  Width           =   795
               End
               Begin VB.Label Label56 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lucro %"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   89
                  Top             =   240
                  Width           =   570
               End
               Begin VB.Label Label55 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor"
                  Height          =   195
                  Left            =   1020
                  TabIndex        =   90
                  Top             =   240
                  Width           =   360
               End
            End
            Begin VB.Frame Frame16 
               Caption         =   "Atac. - À Vista"
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
               TabIndex        =   108
               Top             =   240
               Width           =   1875
               Begin VB.TextBox txtValorAV 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1020
                  TabIndex        =   20
                  Top             =   480
                  Width           =   795
               End
               Begin VB.TextBox txtMargemAV 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   60
                  TabIndex        =   19
                  Top             =   480
                  Width           =   895
               End
               Begin VB.Label Label58 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor"
                  Height          =   195
                  Left            =   1020
                  TabIndex        =   110
                  Top             =   240
                  Width           =   360
               End
               Begin VB.Label Label57 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lucro %"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   109
                  Top             =   240
                  Width           =   570
               End
            End
            Begin VB.Frame Frame17 
               Caption         =   "Atac. - À Prazo"
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
               Left            =   6060
               TabIndex        =   105
               Top             =   240
               Width           =   1875
               Begin VB.TextBox txtMargemAP 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   60
                  TabIndex        =   21
                  Top             =   480
                  Width           =   895
               End
               Begin VB.TextBox txtValorAP 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1020
                  TabIndex        =   22
                  Top             =   480
                  Width           =   795
               End
               Begin VB.Label Label60 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lucro %"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   107
                  Top             =   240
                  Width           =   570
               End
               Begin VB.Label Label59 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor"
                  Height          =   195
                  Left            =   1020
                  TabIndex        =   106
                  Top             =   240
                  Width           =   360
               End
            End
         End
         Begin VB.Frame frmCusto2 
            Caption         =   "Custo"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   98
            Top             =   1080
            Width           =   5115
            Begin VB.OptionButton optUnidUnid 
               Caption         =   "UN > UN"
               Height          =   195
               Left            =   60
               TabIndex        =   138
               Top             =   900
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.OptionButton optPesoMetro 
               Caption         =   "Peso > Metro"
               Height          =   195
               Left            =   2160
               TabIndex        =   137
               Top             =   900
               Width           =   1335
            End
            Begin VB.OptionButton optCXUnid 
               Caption         =   "CX > UN"
               Height          =   195
               Left            =   1140
               TabIndex        =   136
               Top             =   900
               Width           =   1035
            End
            Begin VB.TextBox txtQuantUnid 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   2100
               TabIndex        =   9
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtQuantFinal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Height          =   315
               Left            =   3060
               TabIndex        =   11
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtCusto 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Height          =   315
               Left            =   4080
               TabIndex        =   12
               Top             =   480
               Width           =   915
            End
            Begin VB.TextBox txtValorBruto 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   1140
               TabIndex        =   8
               Top             =   480
               Width           =   915
            End
            Begin VB.TextBox txtQuantBruto 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   60
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   480
               Width           =   1035
            End
            Begin ChamaleonBtn.chameleonButton cmdFracionar 
               Height          =   315
               Left            =   3720
               TabIndex        =   13
               Top             =   840
               Visible         =   0   'False
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Fracionar"
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
               MICON           =   "Entrada_Estoque.frx":00FC
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdUsarUnidMedida 
               Height          =   315
               Left            =   3720
               TabIndex        =   10
               Top             =   840
               Visible         =   0   'False
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Definir Custo"
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
               MICON           =   "Entrada_Estoque.frx":0118
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label39 
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2820
               TabIndex        =   139
               Top             =   420
               Width           =   135
            End
            Begin VB.Label lblFracao 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fração"
               Enabled         =   0   'False
               Height          =   195
               Left            =   2100
               TabIndex        =   103
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Quant. Final"
               Height          =   195
               Left            =   3060
               TabIndex        =   102
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label63 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Unit."
               Height          =   195
               Left            =   4080
               TabIndex        =   101
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Inicio"
               Height          =   195
               Left            =   1140
               TabIndex        =   100
               Top             =   240
               Width           =   780
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Quant. Inicio"
               Height          =   195
               Left            =   60
               TabIndex        =   99
               Top             =   240
               Width           =   900
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdAdicionar2 
            Height          =   315
            Left            =   9060
            TabIndex        =   23
            ToolTipText     =   "Adiciona"
            Top             =   2340
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Adicionar"
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
            MICON           =   "Entrada_Estoque.frx":0134
            PICN            =   "Entrada_Estoque.frx":0150
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdRemover2 
            Height          =   315
            Left            =   10680
            TabIndex        =   25
            ToolTipText     =   "Remove"
            Top             =   2340
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Remover"
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
            MICON           =   "Entrada_Estoque.frx":04EA
            PICN            =   "Entrada_Estoque.frx":0506
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSFlexGridLib.MSFlexGrid GridEntradaItens 
            Height          =   1875
            Left            =   60
            TabIndex        =   24
            Top             =   2700
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   3307
            _Version        =   393216
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin ChamaleonBtn.chameleonButton cmdAtualDesc 
            Height          =   315
            Left            =   1320
            TabIndex        =   122
            ToolTipText     =   "Adiciona"
            Top             =   660
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Atualizar Descrição"
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
            MICON           =   "Entrada_Estoque.frx":08A0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAtualUnid 
            Height          =   315
            Left            =   2880
            TabIndex        =   123
            ToolTipText     =   "Adiciona"
            Top             =   660
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Atualizar Unid. Med."
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
            MICON           =   "Entrada_Estoque.frx":08BC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAtualNCM 
            Height          =   315
            Left            =   4500
            TabIndex        =   124
            ToolTipText     =   "Adiciona"
            Top             =   660
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Atualizar NCM"
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
            MICON           =   "Entrada_Estoque.frx":08D8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAtualProd 
            Height          =   315
            Left            =   8280
            TabIndex        =   125
            ToolTipText     =   "Adiciona"
            Top             =   660
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Cadastrar Produto"
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
            MICON           =   "Entrada_Estoque.frx":08F4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAtualCFOP 
            Height          =   315
            Left            =   5700
            TabIndex        =   131
            ToolTipText     =   "Adiciona"
            Top             =   660
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Atualizar CFOP"
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
            MICON           =   "Entrada_Estoque.frx":0910
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAtualCST 
            Height          =   315
            Left            =   7020
            TabIndex        =   132
            ToolTipText     =   "Adiciona"
            Top             =   660
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Atualizar CST"
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
            MICON           =   "Entrada_Estoque.frx":092C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdLocalizarManual 
            Height          =   315
            Left            =   9720
            TabIndex        =   133
            ToolTipText     =   "Adiciona"
            Top             =   660
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Localizar Manualmente"
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
            MICON           =   "Entrada_Estoque.frx":0948
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdCancelarEntrada 
            Height          =   315
            Left            =   11520
            TabIndex        =   134
            ToolTipText     =   "Adiciona"
            Top             =   660
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
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
            MICON           =   "Entrada_Estoque.frx":0964
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAtualEAN 
            Height          =   315
            Left            =   120
            TabIndex        =   135
            ToolTipText     =   "Adiciona"
            Top             =   660
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Atualizar EAN"
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
            MICON           =   "Entrada_Estoque.frx":0980
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblAviso 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pressione [ F2 ]  para obter o lucro estimado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   5040
            TabIndex        =   127
            Top             =   2340
            Visible         =   0   'False
            Width           =   3960
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblValorItens 
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
            Left            =   13080
            TabIndex        =   126
            Top             =   4560
            Width           =   225
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   3135
         Left            =   120
         TabIndex        =   72
         Top             =   6120
         Width           =   15
         Begin VB.Frame Frame11 
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
            Left            =   10380
            TabIndex        =   88
            Top             =   180
            Width           =   1995
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1080
               TabIndex        =   92
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lucro %"
               Height          =   195
               Left            =   120
               TabIndex        =   91
               Top             =   240
               Width           =   570
            End
         End
         Begin VB.Frame Frame10 
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
            Left            =   8340
            TabIndex        =   84
            Top             =   180
            Width           =   1995
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   120
               TabIndex        =   87
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1080
               TabIndex        =   86
               Top             =   240
               Width           =   360
            End
         End
         Begin VB.Frame Frame9 
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
            Left            =   6300
            TabIndex        =   80
            Top             =   180
            Width           =   1995
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1080
               TabIndex        =   83
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   120
               TabIndex        =   82
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame8 
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
            Left            =   4260
            TabIndex        =   76
            Top             =   180
            Width           =   1995
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem %"
               Height          =   195
               Left            =   120
               TabIndex        =   79
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   1080
               TabIndex        =   78
               Top             =   240
               Width           =   360
            End
         End
         Begin VB.Frame Frame7 
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
            Left            =   720
            TabIndex        =   73
            Top             =   180
            Width           =   1395
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Height          =   315
               Left            =   60
               TabIndex        =   74
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Quant."
               Height          =   195
               Left            =   120
               TabIndex        =   75
               Top             =   240
               Width           =   480
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdAdicionar 
            Height          =   315
            Left            =   0
            TabIndex        =   95
            ToolTipText     =   "Adiciona"
            Top             =   0
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Adicionar"
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
            MICON           =   "Entrada_Estoque.frx":099C
            PICN            =   "Entrada_Estoque.frx":09B8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdRemover 
            Height          =   315
            Left            =   10320
            TabIndex        =   96
            ToolTipText     =   "Remove"
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Remover"
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
            MICON           =   "Entrada_Estoque.frx":0D52
            PICN            =   "Entrada_Estoque.frx":0D6E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
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
            Height          =   4000
            Left            =   200
            TabIndex        =   93
            Top             =   0
            Width           =   4000
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Custo Bruto"
               Height          =   195
               Left            =   60
               TabIndex        =   94
               Top             =   240
               Width           =   825
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   55
         Top             =   8280
         Width           =   15615
         Begin VB.Frame Frame5 
            Caption         =   "Critérios:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   5295
            Begin VB.ComboBox cboOrdem 
               Height          =   315
               Left            =   3360
               TabIndex        =   228
               Top             =   300
               Width           =   1755
            End
            Begin VB.ComboBox cboFiltroNota 
               Height          =   315
               Left            =   660
               TabIndex        =   0
               Top             =   300
               Width           =   1995
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ordem:"
               Height          =   195
               Left            =   2760
               TabIndex        =   229
               Top             =   300
               Width           =   510
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Filtro:"
               Height          =   195
               Left            =   180
               TabIndex        =   68
               Top             =   300
               Width           =   375
            End
         End
         Begin VB.Frame Frame4 
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
            Height          =   735
            Left            =   5520
            TabIndex        =   56
            Top             =   240
            Width           =   6975
            Begin ChamaleonBtn.chameleonButton cmdConNotaCal1 
               Height          =   315
               Left            =   1560
               TabIndex        =   144
               Tag             =   "Calendario"
               Top             =   300
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
               MICON           =   "Entrada_Estoque.frx":1108
               PICN            =   "Entrada_Estoque.frx":1124
               PICH            =   "Entrada_Estoque.frx":3477
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdConNotaCal2 
               Height          =   315
               Left            =   3360
               TabIndex        =   143
               Tag             =   "Calendario"
               Top             =   300
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
               MICON           =   "Entrada_Estoque.frx":57CA
               PICN            =   "Entrada_Estoque.frx":57E6
               PICH            =   "Entrada_Estoque.frx":7B39
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.TextBox txtConNotaCodCliente 
               Height          =   315
               Left            =   4860
               TabIndex        =   58
               Top             =   300
               Visible         =   0   'False
               Width           =   495
            End
            Begin MSMask.MaskEdBox mskConNotaFinal 
               Height          =   315
               Left            =   2460
               TabIndex        =   59
               Top             =   300
               Visible         =   0   'False
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskConNotaInicial 
               Height          =   315
               Left            =   660
               TabIndex        =   60
               Top             =   300
               Visible         =   0   'False
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin VB.ComboBox cboConNotaCliente 
               Height          =   315
               Left            =   1020
               TabIndex        =   70
               Top             =   300
               Visible         =   0   'False
               Width           =   3885
            End
            Begin VB.TextBox txtConNotaNumNota 
               Height          =   315
               Left            =   960
               TabIndex        =   57
               Top             =   300
               Visible         =   0   'False
               Width           =   5895
            End
            Begin VB.ComboBox cboConNotaMes 
               Height          =   315
               Left            =   600
               TabIndex        =   1
               Top             =   300
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.ComboBox cboConNotaAno 
               Height          =   315
               Left            =   2640
               TabIndex        =   2
               Top             =   300
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.Label lblConNotaAno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   2220
               TabIndex        =   66
               Top             =   300
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label lblConNotaMes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   195
               Left            =   180
               TabIndex        =   65
               Top             =   300
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label lblConNotaCliente 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fornecedor:"
               Height          =   195
               Left            =   120
               TabIndex        =   64
               Top             =   300
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblConNotaInicial 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicial:"
               Height          =   195
               Left            =   1980
               TabIndex        =   63
               Top             =   300
               Visible         =   0   'False
               Width           =   450
            End
            Begin VB.Label lblConNotaFinal 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Final:"
               Height          =   195
               Left            =   180
               TabIndex        =   62
               Top             =   300
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label lblConNotaNumNota 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Num. Nota:"
               Height          =   195
               Left            =   120
               TabIndex        =   61
               Top             =   300
               Visible         =   0   'False
               Width           =   810
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirConNotas 
            Height          =   555
            Left            =   12540
            TabIndex        =   69
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "Entrada_Estoque.frx":9E8C
            PICN            =   "Entrada_Estoque.frx":9EA8
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
            Left            =   14040
            TabIndex        =   142
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "Entrada_Estoque.frx":A782
            PICN            =   "Entrada_Estoque.frx":A79E
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
      Begin VB.Frame frmNota 
         Caption         =   "Nota Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   15615
         Begin VB.TextBox txtChave 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   10740
            TabIndex        =   212
            Top             =   540
            Width           =   4455
         End
         Begin VB.TextBox txtTotaldaNota 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   9300
            TabIndex        =   146
            Top             =   540
            Width           =   1395
         End
         Begin VB.TextBox txtCodUsuario 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   11100
            TabIndex        =   140
            TabStop         =   0   'False
            ToolTipText     =   "Item"
            Top             =   0
            Width           =   855
         End
         Begin VB.TextBox txtItem 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   12000
            TabIndex        =   130
            TabStop         =   0   'False
            ToolTipText     =   "Item"
            Top             =   0
            Width           =   855
         End
         Begin VB.TextBox txtCodEntrada 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   12900
            TabIndex        =   128
            TabStop         =   0   'False
            ToolTipText     =   "Cód do Entrada"
            Top             =   0
            Width           =   855
         End
         Begin MSMask.MaskEdBox mskEmissao 
            Height          =   315
            Left            =   7980
            TabIndex        =   6
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtNumNota 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   4
            Top             =   540
            Width           =   1380
         End
         Begin VB.TextBox TxtCodFornecedor 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4740
            MaxLength       =   10
            TabIndex        =   32
            Top             =   300
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.ComboBox cboFornecedor 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1560
            TabIndex        =   5
            Top             =   540
            Width           =   6375
         End
         Begin ChamaleonBtn.chameleonButton cmdCopiarChave 
            Height          =   315
            Left            =   15180
            TabIndex        =   213
            Top             =   540
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "C"
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
            MICON           =   "Entrada_Estoque.frx":C530
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chave"
            Height          =   195
            Left            =   10740
            TabIndex        =   214
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor da Nota"
            Height          =   195
            Left            =   9300
            TabIndex        =   36
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. Emissão"
            Height          =   195
            Left            =   7980
            TabIndex        =   35
            Top             =   300
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº NF"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   300
            Width           =   435
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor"
            Height          =   195
            Left            =   1560
            TabIndex        =   33
            Top             =   300
            Width           =   810
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdNovo 
         Height          =   615
         Left            =   13560
         TabIndex        =   28
         Top             =   6060
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Novo"
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
         MICON           =   "Entrada_Estoque.frx":C54C
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
         Left            =   13560
         TabIndex        =   29
         Top             =   8640
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
         MICON           =   "Entrada_Estoque.frx":C568
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImportarXML 
         Height          =   615
         Left            =   13560
         TabIndex        =   27
         Top             =   5400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Ler XML"
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
         MICON           =   "Entrada_Estoque.frx":C584
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid GridNotas 
         Height          =   7515
         Left            =   -74880
         TabIndex        =   37
         Top             =   480
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   13256
         _Version        =   393216
         TextStyleFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton ChaDownloadXML 
         Height          =   615
         Left            =   13560
         TabIndex        =   3
         Top             =   7380
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Download XML"
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
         MICON           =   "Entrada_Estoque.frx":C5A0
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
         Left            =   13560
         TabIndex        =   26
         Top             =   6720
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
         MICON           =   "Entrada_Estoque.frx":C5BC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdBuscarXML 
         Height          =   615
         Left            =   13560
         TabIndex        =   145
         Top             =   4740
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Baixar XML"
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
         MICON           =   "Entrada_Estoque.frx":C5D8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblSomaNota 
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
         Left            =   -59520
         TabIndex        =   141
         Top             =   8040
         Width           =   225
      End
      Begin VB.Label lblTotalNota 
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
         Left            =   -74880
         TabIndex        =   71
         Top             =   8040
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa"
         Height          =   195
         Left            =   -66300
         TabIndex        =   54
         Top             =   2040
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         Height          =   195
         Left            =   -70575
         TabIndex        =   53
         Top             =   2700
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         Height          =   195
         Left            =   -71685
         TabIndex        =   52
         Top             =   2700
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde Vol."
         Height          =   195
         Left            =   -73560
         TabIndex        =   50
         Top             =   2700
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso bruto"
         Height          =   195
         Left            =   -69570
         TabIndex        =   49
         Top             =   2700
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso liquido"
         Height          =   195
         Left            =   -68625
         TabIndex        =   48
         Top             =   2700
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         Height          =   195
         Left            =   -65445
         TabIndex        =   47
         Top             =   2040
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transportadora"
         Height          =   195
         Left            =   -70320
         TabIndex        =   46
         Top             =   2040
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fatura"
         Height          =   195
         Left            =   -65280
         TabIndex        =   44
         Top             =   1200
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B. ICMS"
         Height          =   195
         Left            =   -67605
         TabIndex        =   43
         Top             =   1440
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B. ICMS Sub."
         Height          =   195
         Left            =   -66405
         TabIndex        =   42
         Top             =   1440
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seguro"
         Height          =   195
         Left            =   -69870
         TabIndex        =   41
         Top             =   1440
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Despesas"
         Height          =   195
         Left            =   -68730
         TabIndex        =   40
         Top             =   1440
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem"
         Height          =   195
         Left            =   -72540
         TabIndex        =   39
         Top             =   1440
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Especie"
         Height          =   195
         Left            =   -72690
         TabIndex        =   51
         Top             =   2700
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do Frete"
         Height          =   195
         Left            =   -64440
         TabIndex        =   45
         Top             =   2040
         Visible         =   0   'False
         Width           =   990
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   38
      Top             =   9660
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21749
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2470
            MinWidth        =   2470
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16:05"
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
End
Attribute VB_Name = "Entrada_Estoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private moCombo As cComboHelper
Public TbEntrada As New ADODB.Recordset
Public TbEntradaItens As New ADODB.Recordset
Public TbProdutos As New ADODB.Recordset
Public TbProdFrac As New ADODB.Recordset
Public TbConsulta As New ADODB.Recordset
Public TbNotaPedido As New ADODB.Recordset
Dim Tb As New ADODB.Recordset
Dim Titulo, Book As Variant, NomeTabela
Dim TbProduto As New ADODB.Recordset
Private iRow As Long, iCol As Long
Dim sSQL As String
Dim r As ADODB.Recordset
Dim Compra As Long
Dim oCfg As ConfigItem

'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1

Private Sub AtualizarGrid_Itens()
  Dim i As Integer
   Dim sSQL As String
   
   For i = 1 To GridNotasItens.rows - 1
      If GridNotasItens.TextMatrix(i, 1) <> "" Then
         dbData.Execute "UPDATE EntradaEstoqueItens SET CFOP = " & GridNotasItens.TextMatrix(i, 5) & ", CST = '" & GridNotasItens.TextMatrix(i, 4) & "', NCM = '" & GridNotasItens.TextMatrix(i, 3) & "'  WHERE CodigoNota = " & txtCodNota.Text & " AND ITEM = " & GridNotasItens.TextMatrix(i, 11) & ""
         dbData.Execute "UPDATE Produtos SET NCM = '" & GridNotasItens.TextMatrix(i, 3) & "'  WHERE CODIGO = " & GridNotasItens.TextMatrix(i, 1) & ""
      End If
   Next
End Sub

Private Sub LimparObjetosProdutoSel()
'txtQuantBruto = ""
'txtValorBruto = ""
txtMargemVV.Text = ""
txtMargemVP.Text = ""
txtMargemAV.Text = ""
txtMargemAP.Text = ""
txtValorVV.Text = ""
txtValorVP.Text = ""
txtValorAV.Text = ""
txtValorAP.Text = ""
txtCusto.Text = ""
txtQuantUnid = ""
txtQuantFinal = ""
End Sub

Sub Load_Data_Itens()
Dim sSQL As String, seq As Integer

If txtCodNota.Text = "" Then Exit Sub

sSQL = "SELECT MAX(Item) r FROM EntradaEstoqueItens WHERE CodigoNota = " & Val(txtCodNota.Text)
seq = SQLExecutaRetorno(sSQL, "r", 0) + 1
Tb("CodigoNota") = Format(txtCodNota.Text, "@")
Tb("Item") = seq
Tb("CodigoProduto") = Format(txtCodProduto, "@")
Tb("NomeProduto") = UCase(Format(cboDescricao, "@"))
Tb("CFOP") = Format(txtCfop, "@")
Tb("NCM") = Format(txtNCM, "@")
Tb("CST") = Right(Format(txtCST, "@"), 3)
Tb("UnidadeComercial") = UCase(Format(txtUnid, "@"))
Tb("ValorUnitarioComercializacao") = CDbl(Format(txtValor, "@"))
Tb("ValorTotalBruto") = CDbl(Format(txtSubtotal, "@"))
Tb("tipodesconto") = Format(1, "@")
Tb("desconto") = CDbl(Format(txtDesc, "@"))
Tb("Valordesconto") = CDbl(Format(txtDesc, "@"))
If txtQuant.Text <> "" Then Tb("QuantidadeComercial") = CDbl(Format(txtQuant, "@")) Else Tb("QuantidadeComercial") = Format(1, "@")
If txtICMS.Text <> "" Then Tb("pICMS") = CDbl(Format(txtICMS, "@"))
If txtICMS.Text <> "" Then Tb("vBC") = CDbl(Format(txtSubtotal, "@"))
If txtICMS.Text <> "" Then Tb("vICMS") = Round(Format(txtSubtotal, "@") * (Format(txtICMS, "@") / 100), 2)
End Sub

Private Sub Calcular_Total()
Dim var_Quant As Double
Dim var_VALOR As Currency, var_Total As Currency

If txtQuant.Text = "" Then var_Quant = 1 Else var_Quant = txtQuant.Text
If txtValor.Text = "" Then var_VALOR = 0 Else var_VALOR = txtValor.Text

var_Total = var_VALOR * var_Quant
txtSubtotal.Text = Format(var_Total, ocMONEY)
End Sub
Sub Clear_Controls()
   Limpa_Tudo Me
End Sub

Private Sub LimparCamposItens()
txtCodBarra.Text = ""
cboDescricao.Text = ""
txtCodProduto.Text = ""
txtCST = Format("0", "@")
txtNCM = Format("0", "@")
txtValor = Format("0", "@")
txtSubtotal = Format("0", "@")
txtQuant = Format("1", "@")
txtICMS = Format("0", "@")
txtCfop = Format("0", "@")
txtDesc = Format("0", "@")
chkDesc.Value = 0
chkDesc.Value = 1
End Sub

Sub Load_Data()
'On Error GoTo erro
    vgDb.BeginTrans
    If Text17.Text = "" Then Text17.Text = "0"
    If Text18.Text = "" Then Text18.Text = "0"
    TbEntrada("CodigoNatureza") = IIf(IsNull(Format(Left(cboDestOperacao.Text, 1), "@")) Or Vazio(Format(Left(cboDestOperacao.Text, 1), "@")), 1, Format(Left(cboDestOperacao.Text, 1), "@"))
    TbEntrada("CodigoNota") = Format(txtCodNota.Text, "@")
    'TbEntrada("SerieNF") = 1
    'TbEntrada("InscricaoEstadual") = 0
    'TbEntrada("Suframa") = 0
    'TbEntrada("CNPJ_CPF") = 0
    'TbEntrada("Logradouro") = 0
    'TbEntrada("numero") = 0
    'TbEntrada("CodigoIBGE") = 0
    'TbEntrada("Bairro") = 0
    'TbEntrada("Complemento") = 0
    'TbEntrada("Municipio") = 0
    'TbEntrada("UF") = 0
    TbEntrada("InformacoesComplementares") = Format(txtInfComple, "@")
    TbEntrada("NaturezaOperacao") = Format(cboNatureza, "@")
    'TbEntrada("TipoDocumento") = 1
    TbEntrada("DataEmissao") = IIf(TbEntrada("DataEmissao") = Empty, Format(Date, "dd/mm/yyyy"), Format(mskEmissao, "@"))
    'TbEntrada("DataSaida") = IIf(TbEntrada("DataSaida") = Empty, Format(Date, "dd/mm/yyyy"), Format(mskSaida, "@"))
    'TbEntrada("HoraSaida") = IIf(TbEntrada("HoraSaida") = Empty, Format(Time(), "HH:MM:ss"), Format(mskHora, "@"))
    
    TbEntrada("TranspCodigo") = IIf(IsNull(Format(txtCodTransporte, "@")) Or Vazio(Format(txtCodTransporte, "@")), 0, Format(txtCodTransporte, "@"))
    TbEntrada("TranspNome") = Format(cboTransporte, "@")
    TbEntrada("TranspPlaca") = Format(txtPlaca, "@")
    TbEntrada("VolumeQuantidade") = Format(Text10, "@")
    TbEntrada("VolumeEspecie") = Format(Text11, "@")
    TbEntrada("VolumeMarca") = Format(Text12, "@")
    TbEntrada("VolumeNumeracao") = Format(txtVolNumeracao, "@")
    TbEntrada("CodigoObservacao") = IIf(IsNull(Format(Text15, "@")) Or Vazio(Format(Text15, "@")), 0, Format(Text15, "@"))
    TbEntrada("ValorFrete") = IIf(Vazio(txtValorFrete), 0, CDbl(Format(txtValorFrete, "##0.00")))
    TbEntrada("ValorSeguro") = IIf(IsNull(Format(Text19, "@")) Or Vazio(Format(Text19, "@")), 0, CDbl(Format(Text19, "##0.00")))
    TbEntrada("ValorOutrasDespesas") = IIf(IsNull(Format(Text16, "@")) Or Vazio(Format(Text16, "@")), 0, CDbl(Format(Text16, "##0.00")))
    TbEntrada("NumeroNota") = Format(txtNumNota, "@")
    TbEntrada("cCodigoNota") = IIf(TbEntrada("cCodigoNota") = 0, GeraCodigoNota, TbEntrada("cCodigoNota"))
    TbEntrada("VolumePesoBruto") = IIf(IsNull(Format(Text17, "@")) Or Vazio(Format(Text17, "@")), 0, CDbl(Format(Text17, "##0.000")))
    TbEntrada("VolumePesoLiquido") = IIf(IsNull(Format(Text18, "@")) Or Vazio(Format(Text18, "@")), 0, CDbl(Format(Text18, "##0.000")))
    TbEntrada("BaseICMS") = IIf(IsNull(Format(Text22, "@")) Or Vazio(Format(Text22, "@")), 0, CDbl(Format(Text22, "##0.000")))
    TbEntrada("BaseICMSST") = IIf(IsNull(Format(Text21, "@")) Or Vazio(Format(Text21, "@")), 0, CDbl(Format(Text21, "##0.00")))
    TbEntrada("TranspPlacaUF") = Format(Text23, "@")
    TbEntrada("CodigoCorrentista") = IIf(IsNull(Format(TxtCodFornecedor, "@")) Or Vazio(Format(TxtCodFornecedor, "@")), 0, Format(TxtCodFornecedor, "@"))
    TbEntrada("RazaoSocial") = Format(cboFornecedor, "@")
    TbEntrada("ModFrete") = IIf(IsNull(Format(Left(cboModFrete.Text, 1), "@")) Or Vazio(Format(Left(cboModFrete.Text, 1), "@")), 9, Format(Left(cboModFrete.Text, 1), "@"))
    TbEntrada("IndicadorFormaPagamento") = Format(cboIndicadorPagamento.Text, "@")
    TbEntrada("FormatoImpressaoDANFE") = Format(cboFormatoDANFe.Text, "@")
    TbEntrada("FormatoEmissaoNFe") = Format(cboTipoEmissao.Text, "@")
    TbEntrada("Cod_Pedido") = Format(txtCodPedido.Text, "@")
    TbEntrada("IdentificadorDestino") = IIf(IsNull(Format(Left(cboDestOperacao.Text, 1), "@")) Or Vazio(Format(Left(cboDestOperacao.Text, 1), "@")), 1, Format(Left(cboDestOperacao.Text, 1), "@"))
    TbEntrada("IndicadorIEDestinatario") = IIf(IsNull(Format(Left(cboTipoContribuinte.Text, 1), "@")) Or Vazio(Format(Left(cboTipoContribuinte.Text, 1), "@")), 1, Format(Left(cboTipoContribuinte.Text, 1), "@"))
    TbEntrada("ConsumidorFinal") = IIf(IsNull(Format(Left(cboConsumidorFinal.Text, 1), "@")) Or Vazio(Format(Left(cboConsumidorFinal.Text, 1), "@")), 1, Format(Left(cboConsumidorFinal.Text, 1), "@"))
   
    'cboConsumidorFinal
    Exit Sub

Resume

'erro:
'    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce"
'    vgDb.RollbackTrans
'    Exit Sub
End Sub
Public Sub Load_Controls()
'On Error GoTo erro
    TxtCodFornecedor = Format(TbEntrada("CodigoCorrentista"), "@")
    cboFornecedor = Format(TbEntrada("RazaoSocial"), "@")
    txtCodEntrada = Format(TbEntrada("cod_entrada"), "@")
    
    'If TbEntrada("CodigoNatureza") = 1 Then
    '    cboDestOperacao.Text = "1 - Operação Interna"
    'ElseIf TbEntrada("CodigoNatureza") = 2 Then
    '    cboDestOperacao.Text = "2 - Operação Interestadual"
    'ElseIf TbEntrada("CodigoNatureza") = 3 Then
    '    cboDestOperacao.Text = "3 - Operação com Exterior"
    'End If
    
    'If TbEntrada("IndicadorIEDestinatario") = 1 Then
    '    cboTipoContribuinte.Text = "1 - CONTRIBUINTE ICMS"
    'ElseIf TbEntrada("IndicadorIEDestinatario") = 2 Then
    '    cboTipoContribuinte.Text = "2 - CONTRIBUINTE ISENTO"
    'ElseIf TbEntrada("IndicadorIEDestinatario") = 9 Then
    '    cboTipoContribuinte.Text = "9 - NÃO CONTRIBUINTE"
    'End If

    'If TbEntrada("ConsumidorFinal") = False Then
    '    cboConsumidorFinal.Text = "0 - NÃO"
    'ElseIf TbEntrada("ConsumidorFinal") = True Then
    '    cboConsumidorFinal.Text = "1 - SIM"
    'End If

'cboConsumidorFinal
    'cboDestOperacao = Format(TbEntrada("CodigoNatureza"), "@")
    txtNumNota = Format(TbEntrada("NumeroNota"), "@")
    txtCodNota = Format(TbEntrada("CodigoNota"), "@")
    
    txtInformacao = Format(TbEntrada("InformacoesAdicionais"), "@")
    cboNatureza = Format(TbEntrada("NaturezaOperacao"), "@")
    mskEmissao = Format(TbEntrada("DataEmissao"), "@")
    mskSaida = Format(TbEntrada("DataSaida"), "@")
    mskHora = Format(TbEntrada("HoraSaida"), "@")

    
    Text15 = Format(TbEntrada("CodigoObservacao"), "@")
    
    txtValorSeguro = Format(TbEntrada("ValorSeguro"), "##,##0.00")
    txtValorOutrasDespesas = Format(TbEntrada("ValorOutrasDespesas"), "##,##0.00")
    txtValorFrete = Format(TbEntrada("ValorFrete"), "##,##0.00")
    txtBaseICMS = Format(TbEntrada("BaseICMS"), "##,##0.00")
    txtValorICMS = Format(TbEntrada("ValorICMS"), "##,##0.00")
    txtBaseICMSST = Format(TbEntrada("BaseICMSST"), "##,##0.00")
    txtValorICMSST = Format(TbEntrada("ValorICMSST"), "##,##0.00")
    txtValorIPI = Format(TbEntrada("ValorIPI"), "##,##0.00")
    txtValorPIS = Format(TbEntrada("ValorPIS"), "##,##0.00")
    txtValorCOFINS = Format(TbEntrada("ValorCOFINS"), "##,##0.00")
    txtValorDesconto = Format(TbEntrada("ValorDesconto"), "##,##0.00")
    txtTotaldosProdutos = Format(TbEntrada("ValorProdutos"), "##,##0.00")
    txtTotaldaNota = Format(TbEntrada("valornota"), "##,##0.00")
    
    'transporte
    If TbEntrada("ModFrete") = 0 Then
        cboModFrete.Text = "0 - POR CONTA DO EMITENTE"
    ElseIf TbEntrada("ModFrete") = 1 Then
        cboModFrete.Text = "1 - POR CONTA DE AMBOS"
    ElseIf TbEntrada("ModFrete") = 2 Then
        cboModFrete.Text = "2 - POR CONTA DO TERCEIROS"
    ElseIf TbEntrada("ModFrete") = 9 Then
        cboModFrete.Text = "9 - SEM FRETE"
    End If
    
    txtCodTransporte = Format(TbEntrada("TranspCodigo"), "@")
    cboTransporte = Format(TbEntrada("TranspNome"), "@")
    txtPlacaUF = Format(TbEntrada("TranspPlacaUF"), "@")
    txtPlaca = Format(TbEntrada("TranspPlaca"), "@")
    txtVolPesoBruto = Format(TbEntrada("VolumePesoBruto"), "@")
    txtVolPesoLiquido = Format(TbEntrada("VolumePesoLiquido"), "@")
    txtVolNumeracao = Format(TbEntrada("VolumeNumeracao"), "@")
    txtVolQuant = Format(TbEntrada("VolumeQuantidade"), "@")
    txtVolEspecie = Format(TbEntrada("VolumeEspecie"), "@")
    txtVolMarca = Format(TbEntrada("VolumeMarca"), "@")

    
    txtChave = Format(TbEntrada("ChavedeAcesso"), "@")
    'Text31 = Format(TbEntrada("NumeroProtocolo"), "@")
    'Text32 = Format(TbEntrada("NumeroRecibo"), "@")
    
    'cboIndicadorPagamento.Text = Format(TbEntrada("IndicadorFormaPagamento"), "@")
    'cboFormatoDANFe.Text = Format(TbEntrada("FormatoImpressaoDANFE"), "@")
    'cboTipoEmissao.Text = Format(TbEntrada("FormatoEmissaoNFe"), "@")
    'txtCodPedido = Format(TbEntrada("cod_pedido"), "@")
    'txtInfComple.Text = Format(TbEntrada("InformacoesComplementares"), "@")

    'habilita/desabilita botões
    cmdImportarXML.Enabled = True
    
    MostraStatus = MostraStatus_F9()

    cmdExcluir.Enabled = True
Exit Sub
Resume

'erro:
'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub Mostrar_ItensNota()
Dim sSQL As String, enviada As Boolean
Dim totalRegistros As Long
    
On Error GoTo ErrLoad

sSQL = "SELECT *, (CASE WHEN adicionada = 1 THEN 'SIM' ELSE 'NÃO' END) AS VarAdicionada FROM EntradaEstoqueItens WHERE CodigoNota = " & Val(txtCodNota.Text)
RsOpen Tb, sSQL

If Tb.RecordCount > 0 Then totalRegistros = Tb.RecordCount
If Tb.RecordCount > 0 Then lblQuant.Caption = Tb.RecordCount

    
'    enviada = SQLExecutaRetorno("SELECT Enviada FROM EntradaEstoque WHERE CodigoNota = " & Val(Frm_NF.txtCodNota.Text), "Enviada", 0)
'
'    If enviada Then
'       cboDestOperacao.Enabled = False
'       Text3.Enabled = False
'       mskEmissao.Enabled = False
'       mskSaida.Enabled = False
'       mskHora.Enabled = False
'       Text7.Enabled = False
'       Text8.Enabled = False
'       txtPlaca.Enabled = False
'       Text10.Enabled = False
'    End If
    
    LimparGridItensNota
    DoEvents
    FormatarGridItensNota Tb
    Exit Sub
    
ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub FormatarGridItensNota(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   
   With GridNotasItens
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 20
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 5300
      .ColWidth(3) = 500
      .ColWidth(4) = 850
      .ColWidth(5) = 640
      .ColWidth(6) = 500
      
      .ColWidth(7) = 530
      .ColWidth(8) = 480
      .ColWidth(9) = 530
      .ColWidth(10) = 750
      .ColWidth(11) = 530
      .ColWidth(12) = 480
      .ColWidth(13) = 530
      
      .ColWidth(14) = 900
      .ColWidth(15) = 700
      .ColWidth(16) = 700
      .ColWidth(17) = 900
      .ColWidth(18) = 0
      .ColWidth(19) = 0
      
      .TextMatrix(0, 1) = "CÓD."
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "UND"
      .TextMatrix(0, 4) = "NCM"
      .TextMatrix(0, 5) = "CFOP"
      .TextMatrix(0, 6) = "CST"
      
      .TextMatrix(0, 7) = "ALIQ"
      .TextMatrix(0, 8) = "PIS"
      .TextMatrix(0, 9) = "ALIQ"
      .TextMatrix(0, 10) = "COFINS"
      .TextMatrix(0, 11) = "ALIQ"
      .TextMatrix(0, 12) = "IPI"
      .TextMatrix(0, 13) = "ALIQ"
      
      .TextMatrix(0, 14) = "VALOR"
      .TextMatrix(0, 15) = "QTDE"
      .TextMatrix(0, 16) = "DESC."
      .TextMatrix(0, 17) = "TOTAL"
      .TextMatrix(0, 18) = "ITEM"
      .TextMatrix(0, 19) = "ADD"

      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      
      .ColAlignment(7) = 1
      .ColAlignment(8) = 1
      .ColAlignment(9) = 1
      .ColAlignment(10) = 1
      .ColAlignment(11) = 1
      .ColAlignment(12) = 1
      .ColAlignment(13) = 1
      
      .ColAlignment(14) = 6
      .ColAlignment(15) = 6
      .ColAlignment(16) = 6
      .ColAlignment(17) = 6
      .ColAlignment(18) = 1
      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = rTabela("CodigoProduto")
            .TextMatrix(.rows - 1, 2) = rTabela("NomeProduto")
            .TextMatrix(.rows - 1, 3) = rTabela("UnidadeComercial")
            .TextMatrix(.rows - 1, 4) = rTabela("NCM")
            .TextMatrix(.rows - 1, 5) = rTabela("CFOP")
            .TextMatrix(.rows - 1, 6) = rTabela("CST")
                        
            .TextMatrix(.rows - 1, 7) = Format(rTabela("pICMS"), ocMONEY)
            .TextMatrix(.rows - 1, 8) = rTabela("pisCST")
            .TextMatrix(.rows - 1, 9) = Format(rTabela("PISpPIS"), ocMONEY)
            .TextMatrix(.rows - 1, 10) = rTabela("cofinsCST")
            .TextMatrix(.rows - 1, 11) = Format(rTabela("COFINSpCOFINS"), ocMONEY)
            .TextMatrix(.rows - 1, 12) = rTabela("IPICST")
            .TextMatrix(.rows - 1, 13) = Format(rTabela("IPIpIPI"), ocMONEY)
            
            .TextMatrix(.rows - 1, 14) = Format(rTabela("ValorUnitarioComercializacao"), ocMONEY)
            .TextMatrix(.rows - 1, 15) = Format(rTabela("QuantidadeComercial"), ocMONEY)
            .TextMatrix(.rows - 1, 16) = Format(rTabela("valordesconto"), ocMONEY)
            .TextMatrix(.rows - 1, 17) = Format(rTabela("ValorTotalBruto"), ocMONEY)
            .TextMatrix(.rows - 1, 18) = rTabela("ITEM")
            .TextMatrix(.rows - 1, 19) = rTabela("varAdicionada")
            
            '.TextMatrix(.Rows - 1, 7) = Format(rTabela("ValorUnitarioComercializacao"), ocMONEY)
            '.TextMatrix(.Rows - 1, 8) = Format(rTabela("QuantidadeComercial"), ocMONEY)
            '.TextMatrix(.Rows - 1, 9) = Format(rTabela("valordesconto"), ocMONEY)
            '.TextMatrix(.Rows - 1, 10) = Format(rTabela("ValorTotalBruto"), ocMONEY)
            '.TextMatrix(.Rows - 1, 11) = rTabela("ITEM")
            '.TextMatrix(.Rows - 1, 12) = rTabela("varAdicionada")
            
            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If
      
    'Deixar vermelho se adicionado
    For i = 1 To .rows - 1
       For j = 0 To .Cols - 1
          .Col = j
          .Row = i
          
          If .TextMatrix(i, 19) = "SIM" Then
             .CellForeColor = vbRed
          Else
             .CellForeColor = vbBlack
          End If
       Next
    Next
    
    'mudar a cor da coluna
    For i = 1 To .rows - 1
       .Row = i
       .Col = 4:   .CellBackColor = &HE0E0E0
       .Col = 5:   .CellBackColor = &HE0E0E0
       .Col = 6:   .CellBackColor = &HE0E0E0
       .Col = 7:   .CellBackColor = &HE0E0E0
       .Col = 8:   .CellBackColor = &HE0E0E0
       .Col = 9:   .CellBackColor = &HE0E0E0
       .Col = 10:   .CellBackColor = &HE0E0E0
       .Col = 11:   .CellBackColor = &HE0E0E0
       .Col = 12:   .CellBackColor = &HE0E0E0
       .Col = 13:   .CellBackColor = &HE0E0E0
    Next

      
      .rows = .rows - 1
      

            
      .Visible = True
      .Redraw = True
      
      SomarGridItens
   End With
End Sub

Private Sub MostrarNotasInicias()
Dim totalRegistros As Long

On Error GoTo ErrLoad

RsOpen TbConsulta, "SELECT top(200) CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 order by CodigoNota desc"

If TbConsulta.RecordCount > 0 Then totalRegistros = TbConsulta.RecordCount

lblTotalNota.Caption = Format(totalRegistros, "00")

LimparGridNotas
FormatarGridNotas TbConsulta

Exit Sub
Resume

ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Set TbConsulta = Nothing
End Sub

Private Sub MostrarProdutoExistente()
optUnidUnid.Value = True
Dim vCodProdFracionado As Integer

If txtCodBarraSel.Text = "" Then
    RsOpen TbProdutos, "SELECT EAN, DESCRICAO, UNID_MEDIDA, QUANT_ESTOQUE, NCM, ICMSCST, CFOP, CODIGO, FRACIONADO, QUANT_FRACAO, CODPROD_FRACAO, ATIVO,   " & _
                   "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
                   "FROM produtos WHERE descricao = '" & txtDescricaoSel.Text & "' AND ATIVO = 1"

    If TbProdutos.EOF And TbProdutos.BOF Then
        lblProdutoCadastrado.Caption = "Esse produto não foi encontrado no seu banco de dados!"
        lblProdutoCadastrado.ForeColor = &HFF&
        txtCodProdExist = ""
        txtCodBarraExist = ""
        txtDescricaoExist = ""
        txtUnidMedExist = ""
        txtQuantExist = ""
        txtValorExist = ""
        txtNCMExist = ""
        txtCSTExist = ""
        txtCFOPExist = ""
        cmdAtualEAN.Enabled = False
        cmdAtualDesc.Enabled = False
        cmdAtualUnid.Enabled = False
        cmdAtualNCM.Enabled = False
        cmdAtualCFOP.Enabled = False
        cmdAtualCST.Enabled = False
        cmdAtualProd.Enabled = True
        cmdLocalizarManual.Enabled = True
        If ShowMsg("O Produto: " & txtDescricaoSel.Text & " não foi encontrado! Você deseja localizar ele manualmente?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then cmdLocalizarManual_Click
        Exit Sub
    Else
        lblProdutoCadastrado.Caption = "Produto Cadastrado"
        lblProdutoCadastrado.ForeColor = &HC00000
        txtCodProdExist = Format(TbProdutos("CODIGO"), "@")
        txtCodBarraExist = Format(TbProdutos("EAN"), "@")
        txtDescricaoExist = Format(TbProdutos("DESCRICAO"), "@")
        txtUnidMedExist = Format(TbProdutos("UNID_MEDIDA"), "@")
        txtQuantExist = Format(TbProdutos("QUANT_ESTOQUE"), "@")
        txtValorExist = Format(TbProdutos("venda"), "##,##0.00")
        txtNCMExist = Format(TbProdutos("NCM"), "@")
        txtCSTExist = Format(TbProdutos("ICMSCST"), "@")
        txtCFOPExist = Format(TbProdutos("CFOP"), "@")
        cmdAtualEAN.Enabled = True
        cmdAtualDesc.Enabled = True
        cmdAtualUnid.Enabled = True
        cmdAtualNCM.Enabled = True
        cmdAtualProd.Enabled = False
        cmdLocalizarManual.Enabled = False
        'If TbProdutos("FRACIONADO") = True Then
        '    MsgBox "produto fracionado1", vbInformation, "Aviso do Sistema"
        '    vCodProdFracionado = TbProdutos("CODPROD_FRACAO")
        '    RsOpen TbProdFrac, "SELECT CODIGO, EAN, DESCRICAO, UNID_MEDIDA, QUANT_ESTOQUE, NCM, ICMSCST, CFOP, CODIGO, FRACIONADO, QUANT_FRACAO, CODPROD_FRACAO,   " & _
        '       "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
        '       "FROM produtos WHERE CODIGO = " & vCodProdFracionado & ""
        '       Debug.Print TbProdFrac
        '    If Not TbProdFrac.EOF And TbProdFrac.BOF Then
        '        lblNomeProdRel.Caption = TbProdFrac("DESCRICAO")
        '    End If
        '    frmProdFracionado.Visible = True
        '    lblNomeProdSel.Caption = txtDescricaoSel.Text
        '
        'Else
        '    MsgBox "produto não fracionado1", vbInformation, "Aviso do Sistema"
        '    'frmProdFracionado.Visible = False
        'End If
'End If
        Exit Sub
    End If
End If

If TbEntradaItens("EAN") = "SEM GTIN" Then
        lblProdutoCadastrado.Caption = "Esse produto não foi encontrado no seu banco de dados!"
        lblProdutoCadastrado.ForeColor = &HFF&
        txtCodProdExist = ""
        txtCodBarraExist = ""
        txtDescricaoExist = ""
        txtUnidMedExist = ""
        txtQuantExist = ""
        txtValorExist = ""
        txtNCMExist = ""
        txtCSTExist = ""
        txtCFOPExist = ""
        cmdAtualEAN.Enabled = False
        cmdAtualDesc.Enabled = False
        cmdAtualUnid.Enabled = False
        cmdAtualNCM.Enabled = False
        cmdAtualCFOP.Enabled = False
        cmdAtualCST.Enabled = False
        cmdAtualProd.Enabled = True
        cmdLocalizarManual.Enabled = True
        If ShowMsg("O Produto: " & txtDescricaoSel.Text & " não foi encontrado! Você deseja localizar ele manualmente?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then cmdLocalizarManual_Click

Else
    RsOpen TbProdutos, "SELECT EAN, DESCRICAO, UNID_MEDIDA, QUANT_ESTOQUE, NCM, ICMSCST, CFOP, CODIGO, FRACIONADO, QUANT_FRACAO, CODPROD_FRACAO, ATIVO,  " & _
                       "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
                       "FROM produtos WHERE COD_BARRA = '" & txtCodBarraSel.Text & "' AND ATIVO = 1"
    
    
    If TbProdutos.EOF And TbProdutos.BOF Then
        lblProdutoCadastrado.Caption = "Esse produto não foi encontrado no seu banco de dados!"
        lblProdutoCadastrado.ForeColor = &HFF&
        txtCodProdExist = ""
        txtCodBarraExist = ""
        txtDescricaoExist = ""
        txtUnidMedExist = ""
        txtQuantExist = ""
        txtValorExist = ""
        txtNCMExist = ""
        txtCSTExist = ""
        txtCFOPExist = ""
        cmdAtualEAN.Enabled = False
        cmdAtualDesc.Enabled = False
        cmdAtualUnid.Enabled = False
        cmdAtualNCM.Enabled = False
        cmdAtualCFOP.Enabled = False
        cmdAtualCST.Enabled = False
        cmdAtualProd.Enabled = True
        cmdLocalizarManual.Enabled = True
        If ShowMsg("O Produto: " & txtDescricaoSel.Text & " não foi encontrado! Você deseja localizar ele manualmente?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then cmdLocalizarManual_Click
    Else
        lblProdutoCadastrado.Caption = "Produto Cadastrado"
        lblProdutoCadastrado.ForeColor = &HC00000
        txtCodProdExist = Format(TbProdutos("CODIGO"), "@")
        txtCodBarraExist = Format(TbProdutos("EAN"), "@")
        txtDescricaoExist = Format(TbProdutos("DESCRICAO"), "@")
        txtUnidMedExist = Format(TbProdutos("UNID_MEDIDA"), "@")
        txtQuantExist = Format(TbProdutos("QUANT_ESTOQUE"), "@")
        txtValorExist = Format(TbProdutos("venda"), "##,##0.00")
        txtNCMExist = Format(TbProdutos("NCM"), "@")
        txtCSTExist = Format(TbProdutos("ICMSCST"), "@")
        txtCFOPExist = Format(TbProdutos("CFOP"), "@")
        cmdAtualEAN.Enabled = True
        cmdAtualDesc.Enabled = True
        cmdAtualUnid.Enabled = True
        cmdAtualNCM.Enabled = True
        cmdAtualCFOP.Enabled = True
        cmdAtualCST.Enabled = True
        cmdAtualProd.Enabled = False
        cmdLocalizarManual.Enabled = False
        'If TbProdutos("FRACIONADO") = True Then
        '    MsgBox "produto fracionado2", vbInformation, "Aviso do Sistema"
        '    vCodProdFracionado = TbProdutos("CODPROD_FRACAO")
        '    RsOpen TbProdFrac, "SELECT CODIGO, EAN, DESCRICAO, UNID_MEDIDA, QUANT_ESTOQUE, NCM, ICMSCST, CFOP, CODIGO, FRACIONADO, QUANT_FRACAO, CODPROD_FRACAO,   " & _
        '       "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
        '       "FROM produtos WHERE CODIGO = " & vCodProdFracionado & ""
        '    'Debug.Print TbProdFrac
        '    If Not TbProdFrac.EOF And TbProdFrac.BOF Then
        '        lblNomeProdRel.Caption = TbProdFrac("DESCRICAO")
        '    End If
        '    frmProdFracionado.Visible = True
        '    lblNomeProdSel.Caption = txtDescricaoSel.Text

        'Else
        '    MsgBox "produto não fracionado2", vbInformation, "Aviso do Sistema"
        '    frmProdFracionado.Visible = False
        'End If
    End If
End If
End Sub

Private Sub PreencherGridNotas()
Dim totalRegistros As Long

On Error GoTo ErrLoad

RsOpen TbConsulta, "SELECT *, " & _
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                "FROM EntradaEstoque order by NumeroNota desc"
                
If TbConsulta.RecordCount > 0 Then totalRegistros = TbConsulta.RecordCount

LimparGridNotas
FormatarGridNotas TbConsulta

Exit Sub
Resume

ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Set TbConsulta = Nothing
End Sub

Private Sub SomarProdutosNota()
Dim sSQL As String, vTotal As Double

On Error GoTo erro

    sSQL = "SELECT ISNULL(SUM(ValorTotalBruto), 0) r FROM EntradaEstoqueItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    vTotal = SQLExecutaRetorno(sSQL, "r", 0)
    
    sSQL = "UPDATE EntradaEstoque SET ValorProdutos = " & FSQL(vTotal, 2) & ", ValorNota = " & FSQL(vTotal, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
    
    Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub

Public Sub CriarEntradaEstoque()
Dim sSQL As String
Dim varData As Date
Dim varHora As Date

varData = Format(Date, "dd/mm/yyyy")
varHora = Format(Time(), "HH:MM:ss")

Dim VarCodEntradaNovo As Long
VarCodEntradaNovo = SQLExecutaRetorno("SELECT ISNULL(MAX(Codigo),0) r FROM produtos_entrada", "r", 0) + 1
sSQL = "INSERT INTO produtos_entrada ( " & _
        "CODIGO, " & _
        "DATA_ENTRADA, " & _
        "HORA_ENTRADA, " & _
        "NOTAFISCAL, " & _
        "VALOR, " & _
        "DATA_EMISSAO, " & _
        "DATA_SAIDA, " & _
        "HORA_SAIDA, " & _
        "COD_FORNECEDOR, " & _
        "COD_TRANSPORTADORA, " & _
        "VALOR_FRETE, " & _
        "Chave, " & _
        "XML, " & _
        "TIPO_FRETE " & _
        " ) " & _
        "SELECT " & VarCodEntradaNovo & ", " & _
        "CONVERT(DATETIME, '" & Format$(varData, ocDATA) & "', 103), " & _
        "CONVERT(DATETIME, '" & Format$(varHora, ocHORA) & "', 103), " & _
        "EntradaEstoque.NumeroNota, " & _
        "EntradaEstoque.ValorNota, " & _
        "EntradaEstoque.DataEmissao, " & _
        "EntradaEstoque.DataSaida, EntradaEstoque.HoraSaida, EntradaEstoque.CodigoCorrentista, EntradaEstoque.TranspCodigo, EntradaEstoque.TranspValorServico, EntradaEstoque.ChavedeAcesso, 1, '0 - EMITENTE' " & _
        "FROM EntradaEstoque " & _
        "WHERE EntradaEstoque.CodigoNota = " & Compra & ";"
dbData.Execute sSQL

txtCodEntrada.Text = VarCodEntradaNovo

dbData.Execute "UPDATE EntradaEstoque SET Cod_entrada = " & txtCodEntrada.Text & "  WHERE CodigoNota = " & Compra & " "
End Sub









Private Sub cboConNotaAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboConNotaAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboConNotaAno.AddItem i
Next

moCombo.AttachTo cboConNotaAno
End Sub


Private Sub cboConNotaCliente_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
'Dim itemAtual As String
'Dim codAtual As String

'itemAtual = cboFornecedor.Text
'codAtual = TxtCodFornecedor.Text
cboConNotaCliente.Clear

sSQL = "SELECT DISTINCT RAZAO, codigo FROM FORNECEDOR ORDER BY RAZAO;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboConNotaCliente.AddItem r("RAZAO")
   cboConNotaCliente.ItemData(cboConNotaCliente.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

'cboFornecedor.Text = itemAtual
'TxtCodFornecedor.Text = codAtual
moCombo.AttachTo cboConNotaCliente
End Sub


Private Sub cboConNotaCliente_LostFocus()
On Error GoTo TrataErro

If cboConNotaCliente.Text = "" Then txtConNotaCodCliente.Text = "": Exit Sub
If cboConNotaCliente.ListIndex = -1 Then txtConNotaCodCliente.Text = "": Exit Sub

txtConNotaCodCliente = cboConNotaCliente.ItemData(cboConNotaCliente.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboConNotaMes_GotFocus()
cboConNotaMes.Clear

cboConNotaMes.AddItem "Janeiro"
cboConNotaMes.AddItem "Fevereiro"
cboConNotaMes.AddItem "Março"
cboConNotaMes.AddItem "Abril"
cboConNotaMes.AddItem "Maio"
cboConNotaMes.AddItem "Junho"
cboConNotaMes.AddItem "Julho"
cboConNotaMes.AddItem "Agosto"
cboConNotaMes.AddItem "Setembro"
cboConNotaMes.AddItem "Outubro"
cboConNotaMes.AddItem "Novembro"
cboConNotaMes.AddItem "Dezembro"

moCombo.AttachTo cboConNotaMes
End Sub






Private Sub cboFiltroNota_Click()
cboFiltroNota_LostFocus
End Sub


Private Sub cboFiltroNota_GotFocus()
Dim vTexto As String
vTexto = cboFiltroNota.Text
cboFiltroNota.Clear
cboFiltroNota.AddItem "TODAS"
cboFiltroNota.AddItem "NUM. NOTA"
cboFiltroNota.AddItem "CHAVE DE ACESSO"
cboFiltroNota.AddItem "FORNECEDOR"
cboFiltroNota.AddItem "PERÍODO"
cboFiltroNota.AddItem "MENSAL"
cboFiltroNota.Text = vTexto

moCombo.AttachTo cboFiltroNota
End Sub

























Private Sub cboFiltroNota_LostFocus()
If cboFiltroNota.Text = "TODAS" Then
    lblConNotaNumNota.Visible = False
    txtConNotaNumNota.Visible = False
    lblConNotaCliente.Visible = False
    cboConNotaCliente.Visible = False
    txtConNotaCodCliente.Visible = False
    lblConNotaAno.Visible = False
    lblConNotaMes.Visible = False
    cboConNotaAno.Visible = False
    cboConNotaMes.Visible = False
    lblConNotaInicial.Visible = False
    lblConNotaFinal.Visible = False
    mskConNotaInicial.Visible = False
    mskConNotaFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
    lblConNotaNumNota.Caption = "Núm. Nota"
    lblConNotaNumNota.Visible = True
    txtConNotaNumNota.Visible = True
    txtConNotaNumNota.Width = 1875
    lblConNotaCliente.Visible = False
    cboConNotaCliente.Visible = False
    txtConNotaCodCliente.Visible = False
    lblConNotaAno.Visible = False
    lblConNotaMes.Visible = False
    cboConNotaAno.Visible = False
    cboConNotaMes.Visible = False
    lblConNotaInicial.Visible = False
    lblConNotaFinal.Visible = False
    mskConNotaInicial.Visible = False
    mskConNotaFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    txtConNotaNumNota.SetFocus
ElseIf cboFiltroNota.Text = "CHAVE DE ACESSO" Then
    lblConNotaNumNota.Caption = "Chave:"
    lblConNotaNumNota.Visible = True
    txtConNotaNumNota.Visible = True
    txtConNotaNumNota.Width = 5895
    lblConNotaCliente.Visible = False
    cboConNotaCliente.Visible = False
    txtConNotaCodCliente.Visible = False
    lblConNotaAno.Visible = False
    lblConNotaMes.Visible = False
    cboConNotaAno.Visible = False
    cboConNotaMes.Visible = False
    lblConNotaInicial.Visible = False
    lblConNotaFinal.Visible = False
    mskConNotaInicial.Visible = False
    mskConNotaFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    txtConNotaNumNota.SetFocus
ElseIf cboFiltroNota.Text = "FORNECEDOR" Then
    lblConNotaCliente.Visible = True
    cboConNotaCliente.Visible = True
    lblConNotaNumNota.Visible = False
    txtConNotaNumNota.Visible = False
    txtConNotaCodCliente.Visible = False
    lblConNotaAno.Visible = False
    lblConNotaMes.Visible = False
    cboConNotaAno.Visible = False
    cboConNotaMes.Visible = False
    lblConNotaInicial.Visible = False
    lblConNotaFinal.Visible = False
    mskConNotaInicial.Visible = False
    mskConNotaFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    cboConNotaCliente.SetFocus
ElseIf cboFiltroNota.Text = "PERÍODO" Then
    lblConNotaInicial.Visible = True
    lblConNotaFinal.Visible = True
    mskConNotaInicial.Visible = True
    mskConNotaFinal.Visible = True
    cmdConNotaCal1.Visible = True
    cmdConNotaCal2.Visible = True
    lblConNotaNumNota.Visible = False
    txtConNotaNumNota.Visible = False
    lblConNotaCliente.Visible = False
    cboConNotaCliente.Visible = False
    txtConNotaCodCliente.Visible = False
    lblConNotaAno.Visible = False
    lblConNotaMes.Visible = False
    cboConNotaAno.Visible = False
    cboConNotaMes.Visible = False
    If mskConNotaInicial.Visible = True Then mskConNotaInicial.SetFocus
ElseIf cboFiltroNota.Text = "MENSAL" Then
    lblConNotaAno.Visible = True
    lblConNotaMes.Visible = True
    cboConNotaAno.Visible = True
    cboConNotaMes.Visible = True
    lblConNotaNumNota.Visible = False
    txtConNotaNumNota.Visible = False
    lblConNotaCliente.Visible = False
    cboConNotaCliente.Visible = False
    txtConNotaCodCliente.Visible = False
    lblConNotaInicial.Visible = False
    lblConNotaFinal.Visible = False
    mskConNotaInicial.Visible = False
    mskConNotaFinal.Visible = False
    cmdConNotaCal1.Visible = False
    cmdConNotaCal2.Visible = False
    If cboConNotaMes.Visible = True Then cboConNotaMes.SetFocus
Else
End If
End Sub

Private Sub cboOrdem_GotFocus()
Dim vTexto As String
vTexto = cboOrdem.Text
cboOrdem.Clear
cboOrdem.AddItem "EMISSÃO"
cboOrdem.AddItem "NUM. NOTA"
cboOrdem.AddItem "FORNECEDOR"
cboOrdem.Text = vTexto
moCombo.AttachTo cboOrdem
End Sub


Private Sub ChaDownloadXML_Click()
 Dim objNFe As snfe.Util, chaveNFe As String, rsEmpresa As New ADODB.Recordset
 Set objNFe = New snfe.Util
 
 RsOpen rsEmpresa, "SELECT * FROM Empresa"
 dirXML = rsEmpresa!DiretorioXML
 
 chaveNFe = InputBox("Informe a chave acesso da NFe ou use o leitor:", "DOWNLOAD XML NFe", "")
 
 If Not Vazio(chaveNFe) Then
    iRetorno = ConfiguraDLLNFeNFCe(55, "1", objNFe)
    iRetorno = objNFe.DownloadXML(chaveNFe, "", msgRetWS, xCaminhoXML)
    'iRetorno = objNFe.DownloadXML(ChaveNFe, rsEmpresa!DiretorioXML, Retira(rsEmpresa!CNPJ, "-./ ", UM_A_UM), msgRetWS, xCaminhoXML) 'antes de mexer
    'msgRetWS = objNFe.DownloadNFe(ChaveNFe, Retira(rsEmpresa!CNPJ, "-./ ", UM_A_UM), rsEmpresa!DiretorioXML)
    If Not iRetorno Then GoTo Caifora
    cStat = objNFe.retDistDFeInt.cStat
    NFeMotivo = objNFe.retDistDFeInt.xMotivo
    
    If cStat = 138 Then
       NFeDataHora = objNFe.retDistDFeInt.dhResp
       If Existe(xCaminhoXML) = True Then ImportarXML (xCaminhoXML)
       CriarEntradaEstoque
    ElseIf cStat = 137 Then
       MsgBox NFeMotivo, vbCritical, vgAtencao
    End If
 End If
 
 Set objNFe = Nothing
 Set rsEmpresa = Nothing
 
 cmdExcluir.Enabled = True
 
 Exit Sub
 
Caifora:
   MsgBox msgRetWS, vbCritical, "ERRO"
   Set objNFe = Nothing
   Set rsEmpresa = Nothing
End Sub

Private Sub chameleonButton1_Click()
txtMargemVP.Text = txtMargemVV.Text
txtMargemAV = txtMargemVV.Text
txtMargemAP = txtMargemVV.Text
CalcularPrecos
End Sub



Private Sub cmdBuscarXML_Click()
ShellExecute hwnd, "open", "https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ+gAVw2g=", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdImprimir_Click()
'colocar o nome da maquina na barra de status
Dim oIni As Ini

Dim var_Impressora As String
Dim r As ADODB.Recordset

'abrindo arquivo .ini

'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
var_Impressora = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

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

Me.Hide

Set r = dbData.OpenRecordset(sSQL)

'If cboConsulta.Text = "TODOS" Or cboConsulta.Text = "NOTA FISCAL" Or cboConsulta.Text = "FORNECEDOR" Or cboConsulta.Text = "MENSAL" Then
   Set REL_Prod_Entrada_NotaXML.Relatorio.Recordset = r
   REL_Prod_Entrada_NotaXML.dfQuant.Caption = lblTotalNota.Caption
   REL_Prod_Entrada_NotaXML.dfBruto.Caption = lblSomaNota.Caption
   
   If cboFiltroNota.Text = "MENSAL" Then
      REL_Prod_Entrada_NotaXML.dfTipo.Caption = "Tipo: Mês = " & cboConNotaMes.Text & "/" & cboConNotaAno.Text
   ElseIf cboFiltroNota.Text = "PERÍODO" Then
      REL_Prod_Entrada_NotaXML.dfTipo.Caption = "Tipo: Datas = " & mskConNotaInicial.Text & " à " & mskConNotaFinal.Text
   ElseIf cboFiltroNota.Text = "FORNECEDOR" Then
      REL_Prod_Entrada_NotaXML.dfTipo.Caption = "Tipo: Fornecedor = " & cboConNotaCliente.Text & ""
   ElseIf cboFiltroNota.Text = "NOTA FISCAL" Then
      REL_Prod_Entrada_NotaXML.dfTipo.Caption = "Tipo: Nota Fiscal Nº " & txtConNotaNumNota.Text & ""
   Else
      REL_Prod_Entrada_NotaXML.dfTipo.Caption = "Tipo: Todas as notas"
   End If
   
   REL_Prod_Entrada_NotaXML.Relatorio.Ativar
   Unload REL_Prod_Entrada_NotaXML
'Else
'     Set REL_Prod_Entrada_PorProduto.Relatorio.Recordset = r
'   REL_Prod_Entrada_PorProduto.dfQuant.Caption = lblQuant.Caption
'   REL_Prod_Entrada_PorProduto.dfBruto.Caption = lblValor.Caption

'   If cboConsulta.Text = "PRODUTO" Then
'      REL_Prod_Entrada_PorProduto.dfTipo.Caption = "Tipo: Produto = " & cboConsDescricao.Text & ""
'   ElseIf cboConsulta.Text = "FABRICANTE" Then
'      REL_Prod_Entrada_PorProduto.dfTipo.Caption = "Tipo: Fabricante = " & cboConsDescricao.Text & ""
'   ElseIf cboConsulta.Text = "REFERÊNCIA" Then
'      REL_Prod_Entrada_PorProduto.dfTipo.Caption = "Tipo: Referência " & cboConsDescricao.Text & ""
'   ElseIf cboConsulta.Text = "CÓD. BARRA" Then
'      REL_Prod_Entrada_PorProduto.dfTipo.Caption = "Tipo: Cód. de Barra " & cboConsDescricao.Text & ""
'   End If
   
'   REL_Prod_Entrada_PorProduto.Relatorio.Ativar
'   Unload REL_Prod_Entrada_PorProduto
'End If

Me.Show 1
End Sub


Private Sub cmdExcluir_Click()
sSQL = "SELECT item FROM EntradaEstoqueItens WHERE (codigonota = " & txtCodNota.Text & ") and adicionada = 1;"
Set r = dbData.OpenRecordset(sSQL)

Dim vQuantProdutos As Integer
vQuantProdutos = r.RecordCount

If r.State <> 0 Then r.Close
Set r = Nothing

If vQuantProdutos < 1 Then
    If ShowMsg("Deseja excluir a entrada da nota: " & txtNumNota.Text & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
    If txtCodNota.Text <> "" Then
        dbData.Execute "DELETE FROM EntradaEstoqueItens WHERE (codigonota = " & txtCodNota.Text & ");"
        dbData.Execute "DELETE FROM EntradaEstoque WHERE (codigonota = " & txtCodNota.Text & ");"
    End If
    If txtCodEntrada.Text <> "" Then
        dbData.Execute "DELETE FROM produtos_entrada WHERE (codigo = " & txtCodEntrada.Text & ");"
        dbData.Execute "DELETE FROM EntradaEstoqueItens WHERE (CODIGO_ENTRADA = " & txtCodEntrada.Text & ");"
    End If
    Limpa_Tudo Me
    cmdExcluir.Enabled = False
    MostrarNotasInicias
    lblQuant.Caption = "00"
Else
    MsgBox "Não é possivel excluir uma nota fiscal ao qual produtos já foram adicionados aos estoque!", vbInformation, "Aviso do Sistema"
End If
End Sub

Private Sub cmdExcluirItem_Click()
If GridNotasItens.TextMatrix(GridNotasItens.Row, 19) = "SIM" Then
    MsgBox "Produto já adicionado no estoque fisico. Não poderá ser excluído da lista!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If GridNotasItens.rows <= 1 Then Exit Sub

If ShowMsg("Deseja remover o produto: " & GridNotasItens.TextMatrix(GridNotasItens.RowSel, 2) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

dbData.Execute "DELETE FROM EntradaEstoqueItens WHERE (codigonota = " & txtCodNota.Text & ") and item = '" & GridNotasItens.TextMatrix(GridNotasItens.RowSel, 18) & "';"

Mostrar_ItensNota
End Sub

Private Sub cmdUsarUnidMedida_Click()
txtCusto = Format(txtValorBruto, "##,##0.00")
txtQuantFinal = txtQuantBruto
End Sub

Private Sub cmdAdicionar2_Click()

If txtCodEntrada.Text = "" Then Exit Sub
If txtCodProdExist.Text = "" Then Exit Sub

If txtQuantFinal.Text = "" Or txtQuantFinal.Text = "0" Then
   ShowMsg "Insira uma quantidade válida!", vbExclamation
   txtQuantFinal.SetFocus
   Exit Sub
End If

If txtCusto.Text = "" Or txtCusto.Text = "0" Then
   ShowMsg "Insira um custo de produto válido!", vbExclamation
   txtCusto.SetFocus
   Exit Sub
End If

Dim var_COD_ITENS As Long

'AUTONUMERAÇÃO
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_entrada_itens;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then var_COD_ITENS = r("cod_itens") + 1

If r.State <> 0 Then r.Close
Set r = Nothing

 Dim varMargemVV As Double
 Dim varMargemVP As Double
 Dim varMargemAV As Double
 Dim varMargemAP As Double
 
 If txtMargemVV.Text = "" Then MsgBox "Margem de lucro em branco!", vbInformation, "Aviso do Sistema": Exit Sub
 If txtMargemVP.Text = "" Then MsgBox "Margem de lucro em branco!", vbInformation, "Aviso do Sistema": Exit Sub
 If txtMargemAV.Text = "" Then MsgBox "Margem de lucro em branco!", vbInformation, "Aviso do Sistema": Exit Sub
 If txtMargemAP.Text = "" Then MsgBox "Margem de lucro em branco!", vbInformation, "Aviso do Sistema": Exit Sub
 
 varMargemVV = Left$(txtMargemVV.Text, Len(txtMargemVV.Text) - 1)
 varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)
 varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)
 varMargemAP = Left$(txtMargemAP.Text, Len(txtMargemAP.Text) - 1)

sSQL = "INSERT INTO produtos_entrada_itens (" & _
   "codigo, " & _
   "codigo_entrada, " & _
   "codigo_produto, " & _
   "descricao, " & _
   "quant, " & _
   "custo, " & _
   "MARGEM_VV, VALOR_VV, MARGEM_VP, VALOR_VP, MARGEM_AV, VALOR_AV, MARGEM_AP, VALOR_AP, ean, itemxml ) VALUES (" & _
   var_COD_ITENS & ", " & txtCodEntrada.Text & ", " & txtCodProdExist.Text & ", '" & txtDescricaoExist.Text & "', " & _
   Replace(CDbl(txtQuantFinal.Text), ",", ".") & ", " & Replace(CCur(txtCusto.Text), ",", ".") & ", " & _
   Replace(CDbl(varMargemVV), ",", ".") & ", " & Replace(CCur(txtValorVV.Text), ",", ".") & ", " & Replace(CDbl(varMargemVP), ",", ".") & ", " & Replace(CCur(txtValorVP.Text), ",", ".") & ", " & Replace(CDbl(varMargemAV), ",", ".") & ", " & Replace(CCur(txtValorAV.Text), ",", ".") & ", " & Replace(CDbl(varMargemAP), ",", ".") & ", " & Replace(CCur(txtValorAP.Text), ",", ".") & ", '" & txtCodBarraExist.Text & "', " & txtItem & ")"

'Adiciona o registro
dbData.Execute sSQL

Preco_Entrada
Quant_Entrada

'Atualiza o estoque do produto
dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque + " & Replace(txtQuantFinal.Text, ",", ".") & " WHERE (codigo = " & txtCodProdExist.Text & ");"

'marcar produto como adicionado ao estoque
dbData.Execute "UPDATE EntradaEstoqueItens SET adicionada = 1 WHERE (codigonota = " & txtCodNota.Text & ") and item = " & txtItem.Text & ";"

Limpar_SubDados
Mostrar_Itens
Mostrar_ItensNota

'txtValor.Text = Format(SomaGrid(GridEntradaItens, 16), "##,##0.000") 'ver depois pq desativei

cmdAtualEAN.Enabled = False
cmdAtualDesc.Enabled = False
cmdAtualUnid.Enabled = False
cmdAtualNCM.Enabled = False
cmdAtualCFOP.Enabled = False
cmdAtualCST.Enabled = False

On Local Error Resume Next
txtCodBarra.SetFocus

End Sub


Private Sub Mostrar_Itens()
Dim sSQL_Itens As String
Dim r As ADODB.Recordset

If txtCodEntrada.Text = "" Then
   sSQL_Itens = "SELECT * FROM produtos_entrada_itens WHERE 1 = 0;"
   Set r = dbData.OpenRecordset(sSQL_Itens)
Else
   sSQL_Itens = "SELECT produtos_entrada_itens.*, produtos_entrada_itens.codigo as varCod,  produtos.COD_BARRA as var_CodBarra, produtos.REF as var_REF, produtos.codigo, produtos.descricao as varDesc, produtos.tamanho, produtos.fabricante, produtos_entrada_itens.CUSTO as var_custo, (produtos_entrada_itens.CUSTO * produtos_entrada_itens.QUANT) as varTotalCustoItem, produtos_entrada_itens.VALOR_VV as var_venda  " & _
          " FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.CODIGO_PRODUTO " & _
          " WHERE (codigo_entrada = " & txtCodEntrada.Text & ") ORDER BY varCod;"
   Set r = dbData.OpenRecordset(sSQL_Itens)
End If

printSQL = sSQL_Itens

FormatarGrid_Itens r
lblValorItens.Caption = FormatNumber(SomaGrid(GridEntradaItens, 16), 2)

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub


Sub FormatarGrid_Itens(rTabela As ADODB.Recordset)
Dim i As Integer, j As Integer
Dim x As Integer

With GridEntradaItens
   .Clear
   .Cols = 18
   .rows = 2
   
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 0
    .ColWidth(3) = 0
    .ColWidth(4) = 0
    .ColWidth(5) = 4500
    .ColWidth(6) = 700
    .ColWidth(7) = 850
    .ColWidth(8) = 650
    .ColWidth(9) = 850
    .ColWidth(10) = 650
    .ColWidth(11) = 850
    .ColWidth(12) = 650
    .ColWidth(13) = 850
    .ColWidth(14) = 650
    .ColWidth(15) = 850
    .ColWidth(16) = 0
    .ColWidth(17) = 500

    .TextMatrix(0, 1) = "COD"
    .TextMatrix(0, 2) = "COD_ENTRADA"
    .TextMatrix(0, 3) = "COD_BARRA"
    .TextMatrix(0, 4) = "CÓD."
    .TextMatrix(0, 5) = "DESCRIÇÃO"
    .TextMatrix(0, 6) = "QTDE"
    .TextMatrix(0, 7) = "CUSTO"
    .TextMatrix(0, 8) = "% VV"
    .TextMatrix(0, 9) = "VALOR"
    .TextMatrix(0, 10) = "% VP "
    .TextMatrix(0, 11) = "VALOR"
    .TextMatrix(0, 12) = "% AV"
    .TextMatrix(0, 13) = "VALOR"
    .TextMatrix(0, 14) = "% AP"
    .TextMatrix(0, 15) = "VALOR"
    .TextMatrix(0, 16) = "SUBTOTAL"
    .TextMatrix(0, 17) = "ITEMXML"
    
   'colocar os cabeçalho em negrito
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   
   'centralizar o titulo
   For j = 0 To .Cols - 1
      .Row = 0
      .Col = j
      .CellAlignment = flexAlignCenterCenter
   Next
   
   .Redraw = False
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = rTabela("varCod")
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("ean"))
            .TextMatrix(.rows - 1, 2) = rTabela("codigo_entrada")
            .TextMatrix(.rows - 1, 4) = Format$(rTabela("codigo_produto"), "0000")
         If tipoEmpresa = 4 Then
            .TextMatrix(.rows - 1, 5) = rTabela("descricao") & " /  " & rTabela("tamanho") & " / " & rTabela("var_ref")
         Else
            .TextMatrix(.rows - 1, 5) = rTabela("descricao")
         End If
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("quant"))
         .TextMatrix(.rows - 1, 7) = Format$(rTabela("custo"), ocMONEY)
         .TextMatrix(.rows - 1, 8) = FormatNumber(rTabela("MARGEM_VV"), 2) & "%"
         .TextMatrix(.rows - 1, 9) = Format$(rTabela("VALOR_VV"), ocMONEY)
         .TextMatrix(.rows - 1, 10) = FormatNumber(rTabela("MARGEM_VP"), 2) & "%"
         .TextMatrix(.rows - 1, 11) = Format$(rTabela("VALOR_VP"), ocMONEY)
         .TextMatrix(.rows - 1, 12) = FormatNumber(rTabela("MARGEM_AV"), 2) & "%"
         .TextMatrix(.rows - 1, 13) = Format$(rTabela("VALOR_AV"), ocMONEY)
         .TextMatrix(.rows - 1, 14) = FormatNumber(rTabela("MARGEM_AP"), 2) & "%"
         .TextMatrix(.rows - 1, 15) = Format$(rTabela("VALOR_AP"), ocMONEY)
         .TextMatrix(.rows - 1, 16) = Format$(rTabela("varTotalCustoItem"), ocMONEY)
         .TextMatrix(.rows - 1, 17) = rTabela("itemxml")
         
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
        'MUDAR COR DE FONTE DA COLUNA
         For i = 1 To .rows - 1
            .Row = i
            .Col = 6
            .CellBackColor = &HC0FFFF
            .CellFontBold = True
         Next
         
        'COLUNA EM NEGRITO
         For i = 1 To .rows - 1
            .Row = i
            .Col = 5
            .CellFontBold = True
         Next
         
        'COLUNA EM NEGRITO
         For i = 1 To .rows - 1
            .Row = i
            .Col = 7
            .CellFontBold = True
         Next
         
        'COLUNA EM NEGRITO
         For i = 1 To .rows - 1
            .Row = i
            .Col = 9
            .CellFontBold = True
         Next
         
        'COLUNA EM NEGRITO
         For i = 1 To .rows - 1
            .Row = i
            .Col = 11
            .CellFontBold = True
         Next
   
   .Redraw = True
   .rows = .rows - 1
End With

End Sub


Private Sub Preco_Entrada()
Dim sSQL As String
Dim r As ADODB.Recordset

'ENTRADA DO PRODUTO
'If cmdSalvar.Enabled = True Then
   Dim AutoNumeracao As Long
   
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
   
   sSQL = "INSERT INTO produtos_precos (Codigo, COD_PRODUTO, Data, FORMA, MARGEM_VV, VALOR_VV, MARGEM_VP, VALOR_VP, MARGEM_AV, VALOR_AV, MARGEM_AP, VALOR_AP, CUSTO, COD_ENTRADA) VALUES (" & _
      AutoNumeracao & ", " & txtCodProdExist.Text & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), 'ENTRADA', " & Replace(CDbl(varMargemVV), ",", ".") & ", " & Replace(CCur(txtValorVV.Text), ",", ".") & ", " & Replace(CDbl(varMargemVP), ",", ".") & ", " & Replace(CCur(txtValorVP.Text), ",", ".") & ", " & Replace(CDbl(varMargemAV), ",", ".") & ", " & Replace(CCur(txtValorAV.Text), ",", ".") & ", " & Replace(CDbl(varMargemAP), ",", ".") & ", " & Replace(CCur(txtValorAP.Text), ",", ".") & ", " & Replace(CCur(txtCusto.Text), ",", ".") & ", " & txtCodEntrada.Text & ");"
   dbData.Execute sSQL
'End If
End Sub

Private Sub Quant_Entrada()
Dim sSQL As String
Dim r As ADODB.Recordset

'ENTRADA DO PRODUTO
'If cmdSalvar.Enabled = True Then
Dim AutoNumeracao As Long

'AUTONUMERAÇÃO
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_quant;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

Dim vCodFunc As Integer
If txtCodUsuario.Text = "" Then
    vCodFunc = 0
Else
    vCodFunc = txtCodUsuario.Text
End If

sSQL = "INSERT INTO produtos_quant (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, QUANT, TIPO, HORA, COD_USUARIO, ESTOQUE) VALUES (" & _
   AutoNumeracao & ", " & txtCodProdExist.Text & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), " & txtCodEntrada.Text & ", 'ENTRADA', " & Replace(CDbl(txtQuantFinal.Text), ",", ".") & ", 'ADIÇÃO', '" & Format(Now, ocHRMN) & "', " & vCodFunc & ", " & Replace(CDbl(txtQuantExist.Text), ",", ".") & ");"
dbData.Execute sSQL
End Sub
Private Sub SomarGridItens()
Dim Total As Currency, SUBTOTAL As Currency, Desc As Currency
Dim i As Integer

SUBTOTAL = 0
Desc = 0
Total = 0

'Sub-Total
With GridNotasItens
   For i = 1 To .rows - 1
      .Col = 0
      .Row = i
      
         .Col = 6
         SUBTOTAL = SUBTOTAL + (.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15))
        '.Col = 10
         Desc = Desc + .TextMatrix(.Row, 16)
         '.Col = 11
         Total = Total + .TextMatrix(.Row, 17)
   Next
End With


lblSubTotal.Caption = Format(SUBTOTAL, ocMONEY)
lblTotalDesc.Caption = Format(Desc, ocMONEY)
lblValorNota.Caption = Format(Total, ocMONEY)
End Sub

Private Function Atualizar_Dados() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE setor SET setor = '" & cboSetor.Text & "' WHERE (cod_setor = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdAtualCFOP_Click()
If ShowMsg("Tem certeza que deseja alterar o CFOP do produto: " & txtCFOPExist.Text & " pelo CFOP " & txtCFOPSel.Text & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "UPDATE produtos SET CFOP = '" & txtCFOPSel.Text & "' WHERE (CODIGO = '" & txtCodProdExist.Text & "');"
    Call MostrarProdutoExistente
End If
End Sub

Private Sub cmdAtualCST_Click()
If ShowMsg("Tem certeza que deseja alterar o CST do produto: " & txtCSTExist.Text & " pelo CST " & txtCSTSel.Text & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "UPDATE produtos SET ICMSCST = '" & txtCSTSel.Text & "' WHERE (CODIGO = '" & txtCodProdExist.Text & "');"
    Call MostrarProdutoExistente
End If
End Sub


Private Sub cmdAtualDesc_Click()
If ShowMsg("Tem certeza que deseja alterar a descrição do produto: " & txtDescricaoExist.Text & " pela descrição " & txtDescricaoSel.Text & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "UPDATE produtos SET DESCRICAO = '" & txtDescricaoSel.Text & "' WHERE (CODIGO = '" & txtCodProdExist.Text & "');"
    Call MostrarProdutoExistente
End If
End Sub

Private Sub cmdAtualEAN_Click()
If ShowMsg("Tem certeza que deseja alterar a EAN do produto: " & txtDescricaoExist.Text & " pela EAN " & txtCodBarraSel.Text & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "UPDATE produtos SET EAN = '" & txtCodBarraSel.Text & "', COD_BARRA = '" & txtCodBarraSel.Text & "' WHERE (CODIGO = '" & txtCodProdExist.Text & "');"
    Call MostrarProdutoExistente
End If
End Sub


Private Sub cmdAtualNCM_Click()
If ShowMsg("Tem certeza que deseja alterar o NCM do produto: " & txtNCMExist.Text & " pelo NCM " & txtNCMSel.Text & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "UPDATE produtos SET NCM = '" & txtNCMSel.Text & "' WHERE (CODIGO = '" & txtCodProdExist.Text & "');"
    Call MostrarProdutoExistente
End If
End Sub

Private Sub cmdAtualProd_Click()
If ShowMsg("Deseja cadastrar o produto " & txtDescricaoSel & " ?", vbInformation + vbYesNo) = vbYes Then
    Load Produtos_Cadastro
    Produtos_Cadastro.SSTab1.Tab = 0
    
    Produtos_Cadastro.CriarNovoProduto
    
    If txtCodBarraSel.Text = "SEM GTIN" Then Produtos_Cadastro.txtCodBarra.Text = "" Else Produtos_Cadastro.txtCodBarra.Text = txtCodBarraSel.Text
    If txtCodBarraSel.Text = "SEM GTIN" Then Produtos_Cadastro.txtEAN.Text = "" Else Produtos_Cadastro.txtEAN.Text = txtCodBarraSel.Text
    Produtos_Cadastro.txtDescricao = TirarEspaco(txtDescricaoSel.Text)
    Produtos_Cadastro.cboUnidMedida = txtUnidMedSel.Text
    Produtos_Cadastro.txtNCM = txtNCMSel.Text
    If Mid(txtCFOPSel, 2, 1) = 4 Then Produtos_Cadastro.txtCfop = "5405" Else Produtos_Cadastro.txtCfop = "5102"
    If Mid(txtCFOPSel, 2, 1) = 4 Then Produtos_Cadastro.txtCST = "500" Else Produtos_Cadastro.txtCST = "102"
    'Produtos_Cadastro.txtCfop = "5102"
    'Produtos_Cadastro.txtCST = "102"
    Produtos_Cadastro.txtCusto = txtValorSel.Text
    Produtos_Cadastro.txtMargemVV = "0"
    Produtos_Cadastro.txtValorVV = "0"
    Produtos_Cadastro.txtMargemVP = "0"
    Produtos_Cadastro.txtValorVP = "0"
    Produtos_Cadastro.txtMargemAV = "0"
    Produtos_Cadastro.txtValorAV = "0"
    Produtos_Cadastro.txtMargemAP = "0"
    Produtos_Cadastro.txtValorAP = "0"
    Produtos_Cadastro.Show
End If
End Sub

Public Function TirarEspaco(ByVal Value As String) As String
Dim bRepete As Boolean
Value = Replace$(Value, "'", vbNullString)
Do
  Value = Replace$(Value, "  ", " ")
  bRepete = InStr(1, Value, "  ", vbTextCompare)
  Value = Trim(Value)
Loop Until Not bRepete

TirarEspaco = Value
End Function

Private Sub cmdAtualUnid_Click()
If ShowMsg("Tem certeza que deseja alterar a Unid. de Medida do produto: " & txtUnidMedExist.Text & " pela Unid. Medida " & txtUnidMedSel.Text & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "UPDATE produtos SET UNID_MEDIDA = '" & txtUnidMedSel.Text & "' WHERE (CODIGO = '" & txtCodProdExist.Text & "');"
    Call MostrarProdutoExistente
End If
End Sub

Private Sub MostrarValorVenda()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vrVenda As Currency
If txtCodProduto.Text = "" Then Exit Sub

'mostrar o ultimo preço de compra
sSQL = "SELECT TOP 1 venda FROM produtos_entrada_itens WHERE (codigo_produto = " & txtCodProduto & ") ORDER BY codigo DESC;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then vrVenda = r("venda")
If r.State <> 0 Then r.Close
Set r = Nothing

txtValor.Text = Format(vrVenda, ocMONEY)
End Sub

Private Sub cmdCancelarEntrada_Click()
lblProdutoCadastrado.Caption = "Esse produto não foi encontrado no seu banco de dados!"
lblProdutoCadastrado.ForeColor = &HFF&
txtCodProdExist = ""
txtCodBarraExist = ""
txtDescricaoExist = ""
txtUnidMedExist = ""
txtQuantExist = ""
txtValorExist = ""
txtNCMExist = ""
txtCSTExist = ""
txtCFOPExist = ""
cmdAtualEAN.Enabled = False
cmdAtualDesc.Enabled = False
cmdAtualUnid.Enabled = False
cmdAtualNCM.Enabled = False
cmdAtualCFOP.Enabled = False
cmdAtualCST.Enabled = False
cmdAtualProd.Enabled = True
cmdLocalizarManual.Enabled = True
Exit Sub
End Sub

Private Sub cmdConNotaCal1_Click()
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

mskConNotaInicial = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdConNotaCal2_Click()
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

mskConNotaFinal = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub


Private Sub cmdExibirConNotas_Click()
Dim totalRegistros As Long

On Error GoTo ErrLoad

'INDICE
Dim INDICE As String
If cboOrdem.Text = "EMISSÃO" Then
   INDICE = "DataEmissao DESC;"
ElseIf cboOrdem.Text = "NUM. NOTA" Then
   INDICE = "NumeroNota;"
ElseIf cboOrdem.Text = "FORNECEDOR" Then
   INDICE = "RazaoSocial;"
Else
   INDICE = "DataEmissao DESC;"
End If

If cboFiltroNota.Text = "TODAS" Then
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 " & _
                    "ORDER BY " & INDICE
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                        "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                        "FROM EntradaEstoque AS EntradaEstoque_1 " & _
                        "ORDER BY " & INDICE
ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
    If txtConNotaNumNota.Text = "" Then Exit Sub
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 WHERE NumeroNota = " & txtConNotaNumNota & " " & _
                    "ORDER BY " & INDICE
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                        "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                        "FROM EntradaEstoque AS EntradaEstoque_1 WHERE NumeroNota = " & txtConNotaNumNota & " " & _
                        "ORDER BY " & INDICE
ElseIf cboFiltroNota.Text = "CHAVE DE ACESSO" Then
    If txtConNotaNumNota.Text = "" Then Exit Sub
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, ChavedeAcesso," & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 WHERE ChavedeAcesso = '" & txtConNotaNumNota & "' " & _
                    "ORDER BY " & INDICE
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, ChavedeAcesso, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                        "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                        "FROM EntradaEstoque AS EntradaEstoque_1 WHERE ChavedeAcesso = '" & txtConNotaNumNota & "' " & _
                        "ORDER BY " & INDICE
                        'Debug.Print sSQL
ElseIf cboFiltroNota.Text = "FORNECEDOR" Then
    If txtConNotaCodCliente.Text = "" Then Exit Sub
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 WHERE CodigoCorrentista = " & txtConNotaCodCliente.Text & " " & _
                    "ORDER BY " & INDICE
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                        "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                        "FROM EntradaEstoque AS EntradaEstoque_1 WHERE CodigoCorrentista = " & txtConNotaCodCliente.Text & " " & _
                        "ORDER BY " & INDICE
ElseIf cboFiltroNota.Text = "PERÍODO" Then
    If mskConNotaInicial.Text = "" Then Exit Sub
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 WHERE (DataEmissao >= CONVERT(DATETIME, '" & Format(mskConNotaInicial.Text, ocDATA) & "', 103)) AND (DataEmissao <= CONVERT(DATETIME, '" & Format(mskConNotaFinal.Text, ocDATA) & "', 103)) " & _
                    "ORDER BY " & INDICE
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                        "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                        "FROM EntradaEstoque AS EntradaEstoque_1 WHERE (DataEmissao >= CONVERT(DATETIME, '" & Format(mskConNotaInicial.Text, ocDATA) & "', 103)) AND (DataEmissao <= CONVERT(DATETIME, '" & Format(mskConNotaFinal.Text, ocDATA) & "', 103)) " & _
                        "ORDER BY " & INDICE
ElseIf cboFiltroNota.Text = "MENSAL" Then
    If cboConNotaMes.Text = "" Then Exit Sub
    
    Dim vIndMes As Integer
    Dim vWhere As String

If cboConNotaMes.ListCount = 0 Then
    If cboConNotaMes.Text = "janeiro" Then
        vIndMes = cboConNotaMes.ListIndex + 2
    ElseIf cboConNotaMes.Text = "fevereiro" Then
        vIndMes = cboConNotaMes.ListIndex + 3
    ElseIf cboConNotaMes.Text = "março" Then
        vIndMes = cboConNotaMes.ListIndex + 4
    ElseIf cboConNotaMes.Text = "abril" Then
        vIndMes = cboConNotaMes.ListIndex + 5
    ElseIf cboConNotaMes.Text = "maio" Then
        vIndMes = cboConNotaMes.ListIndex + 6
    ElseIf cboConNotaMes.Text = "junho" Then
        vIndMes = cboConNotaMes.ListIndex + 7
    ElseIf cboConNotaMes.Text = "julho" Then
        vIndMes = cboConNotaMes.ListIndex + 8
    ElseIf cboConNotaMes.Text = "agosto" Then
        vIndMes = cboConNotaMes.ListIndex + 9
    ElseIf cboConNotaMes.Text = "setembro" Then
        vIndMes = cboConNotaMes.ListIndex + 10
    ElseIf cboConNotaMes.Text = "outubro" Then
        vIndMes = cboConNotaMes.ListIndex + 11
    ElseIf cboConNotaMes.Text = "novembro" Then
        vIndMes = cboConNotaMes.ListIndex + 12
    ElseIf cboConNotaMes.Text = "dezembro" Then
        vIndMes = cboConNotaMes.ListIndex + 13
    End If
    
    vWhere = "(MONTH(DataEmissao) = " & vIndMes & ") "
Else
    vWhere = "(MONTH(DataEmissao) = " & cboConNotaMes.ListIndex + 1 & ") "
End If

    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 WHERE " & vWhere & " AND (YEAR(DataEmissao) = " & cboConNotaAno & ") order by CodigoNota desc"
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDÊNCIA' END) AS SITUACAO,  " & _
                        "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                        "FROM EntradaEstoque AS EntradaEstoque_1 WHERE " & vWhere & " AND (YEAR(DataEmissao) = " & cboConNotaAno & ") order by CodigoNota desc"
End If

'Debug.Print sSQL

If TbConsulta.RecordCount > 0 Then totalRegistros = TbConsulta.RecordCount
lblTotalNota.Caption = Format(totalRegistros, "00")

LimparGridNotas
FormatarGridNotas TbConsulta

Exit Sub
Resume

ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Set TbConsulta = Nothing
End Sub

Private Sub cmdFracionar_Click()
If txtQuantBruto.Text = "" Then Exit Sub
If txtQuantUnid.Text = "" Then Exit Sub
Dim varQuantBruta As Double
Dim varQuantFrancionada As Double
Dim varQuantResultado As Double
Dim varValorBruto As Currency
    
If optCXUnid.Value = True Then
    varQuantBruta = txtQuantBruto
    varQuantFrancionada = txtQuantUnid
    varQuantResultado = varQuantBruta * varQuantFrancionada
    txtQuantFinal = varQuantResultado
    
    varValorBruto = txtValorBruto.Text
    txtCusto.Text = Format(varValorBruto / varQuantFrancionada, "##,##0.00")
ElseIf optPesoMetro.Value = True Then
    varQuantBruta = txtQuantBruto
    varQuantFrancionada = txtQuantUnid
    varQuantResultado = varQuantBruta * varQuantFrancionada
    txtQuantFinal = varQuantResultado
    
    varValorBruto = txtValorBruto.Text
    txtCusto.Text = Format(varValorBruto / varQuantResultado, "##,##0.00")
End If
End Sub



Private Sub cmdLocalizarManual_Click()
Produtos_LocalizarManual.Show
Produtos_LocalizarManual.lblProduto.Caption = txtDescricaoSel.Text
End Sub

Private Sub cmdNovo_Click()
Limpa_Tudo Me
cmdExcluir.Enabled = False
End Sub

Private Sub cmdRemover2_Click()
If GridEntradaItens.rows <= 1 Then Exit Sub

If GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 1) <> "" Then
    'verificar se o nfce foi gravado
    'Dim sSQL2 As String
    'Dim r2 As ADODB.Recordset

    'sSQL2 = "SELECT EAN FROM produtos_entrada_itens WHERE (codigo_entrada = " & txtCodEntrada.Text & ") AND (CODIGO_PRODUTO = " & GridEntradaItens.TextMatrix(GridEntradaItens.Row, 4) & ");"
    'Set r2 = dbData.OpenRecordset(sSQL2)
    
    'If r2("EAN") <> "" Then
    '    dbData.Execute "UPDATE EntradaEstoqueItens SET adicionada = 0 WHERE (codigonota = " & txtCodNota.Text & ") and item = '" & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 3) & "';"
    'Else
        dbData.Execute "UPDATE EntradaEstoqueItens SET adicionada = 0 WHERE (codigonota = " & txtCodNota.Text & ") and item = '" & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 17) & "';"
    'End If
    
    'If r2.State <> 0 Then r2.Close
    'Set r2 = Nothing
    
    dbData.Execute "DELETE FROM produtos_entrada_itens WHERE (codigo = " & GridEntradaItens.TextMatrix(GridEntradaItens.Row, 1) & ") AND (codigo_entrada = " & txtCodEntrada.Text & ");"
    dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 6) & " WHERE (codigo = " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 4) & ");"
    dbData.Execute "DELETE FROM Produtos_Quant WHERE (COD_PRODUTO = " & GridEntradaItens.TextMatrix(GridEntradaItens.Row, 4) & ") AND (COD_ENTRADA = " & txtCodEntrada.Text & ");"
    dbData.Execute "DELETE FROM Produtos_precos WHERE (COD_PRODUTO = " & GridEntradaItens.TextMatrix(GridEntradaItens.Row, 4) & ") AND (COD_ENTRADA = " & txtCodEntrada.Text & ");"
End If

Mostrar_Itens
Mostrar_ItensNota
End Sub


Private Sub cmdSair_Click()
Unload Me
'CriarEntradaEstoque
End Sub

Private Sub Limpar_SubDados()
txtCodProdExist = ""
txtCodBarraExist = ""
txtDescricaoExist = ""
txtUnidMedExist = ""
txtQuantExist = ""
txtValorExist = ""
txtNCMExist = ""
txtCSTExist = ""
txtCFOPExist = ""

txtCodBarraSel = ""
txtDescricaoSel = ""
txtUnidMedSel = ""
txtQuantSel = ""
txtValorSel = ""
txtNCMSel = ""
txtCSTSel = ""
txtCFOPSel = ""
txtQuantBruto = ""
txtValorBruto = ""
txtMargemVV.Text = ""
txtMargemVP.Text = ""
txtMargemAV.Text = ""
txtMargemAP.Text = ""
txtValorVV.Text = ""
txtValorVP.Text = ""
txtValorAV.Text = ""
txtValorAP.Text = ""
txtCusto.Text = ""
txtQuantUnid = ""
txtQuantFinal = ""

lblProdutoCadastrado.Caption = ""
txtItem.Text = ""
End Sub


Private Sub cmdImportarXML_Click()
If txtCodEntrada.Text <> "" Or txtNumNota.Text <> "" Then: MsgBox "Existe uma nota fiscal(xml) sendo usada!", vbInformation, "Aviso do Sistema": Exit Sub

Dim i As Integer, vgRetVal As String, dlgArquivo As New cCommonDialog
Dim vgArqAtual As String, arqXML As String

vgArqAtual$ = App.path
vgRetVal$ = vgArqAtual$

On Error GoTo deuErro

With dlgArquivo                                  'vamos apresentar dilogo para pegar arquivo
 .CancelError = True                             'vamos testar se o usurio cancelou
 i = Rat(vgArqAtual$, "\")                       'tem diretorio
 If i > 0 Then                                   'coloca
  .InitDir = Left$(vgArqAtual$, i)               'o diretrio default
 End If
 .DialogTitle = "Selecionar o arquivo XML para importação"  'tiulo do dilogo...
 .Filter = "Selecione o XML a Ser Importado|*.xml|Todos os Arquivos (*.*)|*.*"          'filtra todos
 .FilterIndex = 0                                'vamos oferecer o 1 primeiro
 .Flags = OFN_CREATEPROMPT                       'se nao existir, poderemos criar?
 .ShowOpen
 vgRetVal$ = .FileName                           'arquivo a capturado
 If Existe(vgRetVal$) Then
    If vgRetVal$ = "" Then
        Exit Sub
    Else
        arqXML = vgRetVal$
        Call ImportarXML(arqXML)
        CriarEntradaEstoque
    End If
 End If
End With
cmdExcluir.Enabled = True
Exit Sub

deuErro:
 Set dlgArquivo = Nothing
 MsgBox Err.Description, vbCritical, vgAtencao
 Err.Clear
End Sub

Public Sub ImportarXML(arqXML As String)
Dim CNPJ As String, CPF As String, tipoEmp As Integer
Dim cstpis As String, cstcofins As String, csticms As String, cstipi As String, cEAN As String
Dim i As Integer
Dim qtdProd As Integer, qtdDup As Integer, xCfop As Integer, qtfracao As Integer
Dim ICMSBase As Double, ICMSAliquota As Double, ICMSValor As Double
Dim CodProduto As Long, Codigo_fornecedor As Long
Dim Fornecedores As New ADODB.Recordset, Compras_XML As New ADODB.Recordset, TBVerificarNFe As New ADODB.Recordset, Compras_Itens_XML As New ADODB.Recordset, Produtos As New ADODB.Recordset
Dim Produtos_fornecedores As New ADODB.Recordset

Dim XMLDOC As MSXML2.DOMDocument30
Set XMLDOC = New MSXML2.DOMDocument30
XMLDOC.Load (arqXML)

Dim str_sql As String
Dim str_sql_item As String

On Error Resume Next

'-------- IMPORTANDO OS DADOS DO FORNECEDOR ---------

'Trata da Máscara Cnpj\cpf
'CNPJ = Retira(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/CNPJ").Item(0).Text, "./-", UM_A_UM) 'desmarquei
CNPJ = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/CNPJ").Item(0).Text

If Len(CNPJ) < 14 Then
   CPF = Mid(CNPJ, 1, 3) & "." & Mid(CNPJ, 4, 3) & "." & Mid(CNPJ, 7, 3) & "-" & Mid(CNPJ, 10, 2) 'desmarquei
   'CPF = CNPJ
   CNPJ = ""
   tipoEmp = 0
   'Monta String SQL
   str_sql = "Select * from fornecedor where CPF = '" & CPF & "'"
ElseIf Len(CNPJ) = 14 Then
   CNPJ = Mid(CNPJ, 1, 2) & "." & Mid(CNPJ, 3, 3) & "." & Mid(CNPJ, 6, 3) & "/" & Mid(CNPJ, 9, 4) & "-" & Mid(CNPJ, 13, 2) 'desmarquei
   'CNPJ = CNPJ
   tipoEmp = 1
   'Monta String SQL
   str_sql = "Select * from fornecedor where CPF = '" & CNPJ & "'"
End If

'Filta no BD
RsOpen Fornecedores, str_sql

   'Set Fornecedores = dbData.OpenRecordset(str_sql)
   
   
   'If Fornecedores.State <> 0 Then Fornecedores.Close
   'Set Fornecedores = Nothing



'Verifica se o Fornecedor está cadastrado ou não
If Fornecedores.RecordCount = 0 Then
    Dim varCodFornecedor As Integer
    Dim varRAZAOForn As String
    Dim varFantasiaFor As String
    Dim varCNPJforn As String
    Dim varIEforn As String
    Dim varENDERECOforn As String
    Dim varNUMforn As Integer
    Dim varBairroforn As String
    Dim varCidadeforn As String
    Dim varEstadoforn As String
    Dim varCEPform As String
    Dim varCODIBGEform As Integer
    
    varCodFornecedor = SQLExecutaRetorno("SELECT ISNULL(MAX(Codigo),0) r FROM fornecedor", "r", 0) + 1
    varRAZAOForn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xNome").Item(0).Text), 35)
    varFantasiaFor = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xFant").Item(0).Text), 35)
    If tipoEmp = 0 Then varCNPJforn = CPF
    If tipoEmp = 1 Then varCNPJforn = CNPJ
    varIEforn = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/IE").Item(0).Text
    varENDERECOforn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xLgr").Item(0).Text), 40) & ", " & Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").Item(0).Text), 8)
    varBairroforn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xBairro").Item(0).Text), 20)
    varCidadeforn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xMun").Item(0).Text), 20)
    varEstadoforn = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").Item(0).Text)
    varCEPform = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/CEP").Item(0).Text
    varNUMforn = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").Item(0).Text
    varCODIBGEform = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/cMun").Item(0).Text
    
    
    'Comando de inclusão
    sSQL = "INSERT INTO fornecedor (" & _
       "codigo, " & _
       "RAZAO, " & _
       "Fantasia, " & _
       "CPF, " & _
       "IE, " & _
       "ENDERECO, " & _
       "bairro, " & _
       "Cidade, " & _
       "Estado, " & _
       "CEP, " & _
       "Numero, " & _
       "CodigoIBGE ) VALUES ( "
    
    sSQL = sSQL & _
       varCodFornecedor & ", '" & varRAZAOForn & "', '" & varFantasiaFor & "', '" & varCNPJforn & "', '" & _
       varIEforn & "', '" & varENDERECOforn & "', '" & varBairroforn & "', '" & varCidadeforn & "', '" & _
       varEstadoforn & "', '" & varCEPform & "', " & varNUMforn & ", " & varCODIBGEform & ")"
    
    dbData.Execute (sSQL)

Else
   Codigo_fornecedor = Fornecedores!Codigo
   Fornecedores!Endereco = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xLgr").Item(0).Text), 40) & ", " & Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").Item(0).Text), 8)
   'Fornecedores!Num = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").Item(0).Text), 8)
   Fornecedores!bairro = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xBairro").Item(0).Text), 20)
   'Fornecedores!Codigo_municipio = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/cMun").Item(0).Text
   Fornecedores!Cidade = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xMun").Item(0).Text), 20)
   Fornecedores!Estado = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").Item(0).Text)
   Fornecedores!CEP = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/CEP").Item(0).Text
   'Fornecedores!Cod_pais = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/cPais").Item(0).Text
   'Fornecedores!CodigoPais = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/cPais").Item(0).Text
   'Fornecedores!NomePais = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xPais").Item(0).Text
   Fornecedores.Update
End If
'-------- FIM DA IMPORTAÇÃO DO FORNECEDOR --------

   '--------CADASTRO DA ENTRADA NO ESTOQUE ---------
   RsOpen Compras_XML, "SELECT TOP 1 * FROM EntradaEstoque"
   
   Dim vNumNota As Long
   Dim vChave As String
   vNumNota = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/nNF").Item(0).Text
   vChave = XMLDOC.getElementsByTagName("nfeProc/protNFe/infProt/chNFe").Item(0).Text
   
   'sSQL = "SELECT ChavedeAcesso " & _
                "FROM EntradaEstoque WHERE ChavedeAcesso = '" & vChave
   sSQL = "SELECT NumeroNota " & _
                "FROM EntradaEstoque WHERE NumeroNota = " & vNumNota
   Set r = dbData.OpenRecordset(sSQL)
    
    If r.RecordCount <> 0 Then
        MsgBox "Essa Nota Fiscal já foi adicionada anteriormente!", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If
   
   Compra = SQLExecutaRetorno("SELECT ISNULL(MAX(CodigoNota),0) r FROM EntradaEstoque", "r", 0) + 1
   
   Compras_XML.AddNew
   Compras_XML!CodigoNota = Compra
   Compras_XML!ModeloNF = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/mod").Item(0).Text
   Compras_XML!NumeroNota = Format(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/nNF").Item(0).Text, "000000000")
   Compras_XML!SerieNF = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/serie").Item(0).Text
   Compras_XML!ChavedeAcesso = XMLDOC.getElementsByTagName("nfeProc/protNFe/infProt/chNFe").Item(0).Text
   Compras_XML!NaturezaOperacao = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/natOp").Item(0).Text)
   
   If XMLDOC.getElementsByTagName("nfeProc/protNFe").Item(0).Attributes.getNamedItem("versao").Text = "2.00" Then
       Compras_XML!DataEmissao = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/dEmi").Item(0).Text
       Compras_XML!DataSaida = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/dEmi").Item(0).Text
   Else
       Compras_XML!DataEmissao = Left(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/dhEmi").Item(0).Text, 10)
       Compras_XML!DataSaida = Left(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/dhEmi").Item(0).Text, 10)
   End If
  
   'Compras_XML!TipoEntradaEst = "Nota Fiscal"
   Compras_XML!HoraSaida = Time
   
   'fornecedor
   If Fornecedores.RecordCount = 0 Then
       Compras_XML!CodigoCorrentista = varCodFornecedor
    Else
       Compras_XML!CodigoCorrentista = Codigo_fornecedor
    End If
       
   Compras_XML!RazaoSocial = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xNome").Item(0).Text)
   Compras_XML!InscricaoEstadual = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/IE").Item(0).Text)
   
   'Compras_XML!CNPJ_CPF = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/CNPJ").Item(0).Text)
   If tipoEmp = 0 Then Compras_XML!CNPJ_CPF = CPF
   If tipoEmp = 1 Then Compras_XML!CNPJ_CPF = CNPJ
   
   Compras_XML!logradouro = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xLgr").Item(0).Text)
   Compras_XML!Numero = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").Item(0).Text)
   'Compras_XML!Complemento = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xCpl").Item(0).Text)
   Compras_XML!bairro = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xBairro").Item(0).Text)
   Compras_XML!CodigoIBGE = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/cMun").Item(0).Text)
   Compras_XML!Municipio = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xMun").Item(0).Text)
   Compras_XML!UF = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").Item(0).Text)
   Compras_XML!CEP = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/CEP").Item(0).Text)
   Compras_XML!Telefone = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/fone").Item(0).Text)

   'Select Case Int(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/indPag").Item(0).Text)
   '   Case 0
          Compras_XML!IndicadorFormaPagamento = "0 - Pagamento à vista"
   '   Case 1
   '       Compras_XML!IndicadorFormaPagamento = "1 - Pagamento à prazo"
   '   Case 2
   '       Compras_XML!IndicadorFormaPagamento = "2 - Outros"
   '   Case Else
   '       Compras_XML!IndicadorFormaPagamento = "2 - Outros"
   'End Select
   
   Select Case Int(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/tpImp").Item(0).Text)
      Case 1
          Compras_XML!FormatoImpressaoDANFE = "1 - Retrato"
      Case 2
          Compras_XML!FormatoImpressaoDANFE = "2 - Paisagem"
   End Select
   
   Select Case Int(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/tpEmis").Item(0).Text)
      Case 1
          Compras_XML!FormatoEmissaoNFe = "1 - Normal"
      Case 2
          Compras_XML!FormatoEmissaoNFe = "2 - Contingência FS"
      Case 3
          Compras_XML!FormatoEmissaoNFe = "3 - Contingência SCAN"
      Case 4
          Compras_XML!FormatoEmissaoNFe = "4 - Contingência DPEC"
      Case 5
          Compras_XML!FormatoEmissaoNFe = "5 - Contingência FS-DA"
   End Select
   
   Compras_XML!CodigoNatureza = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/idDest").Item(0).Text
   Compras_XML!IndicadorPresencaComprador = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/indPres").Item(0).Text
   Compras_XML!IndicadorIEDestinatario = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/indPres").Item(0).Text
   Compras_XML!FinalidadeEmissaoNFe = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/dest/indIEDest").Item(0).Text
   
   Compras_XML!modFrete = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/transp/modFrete").Item(0).Text
   
   If Int(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/infAdic/infCpl").Length) > 0 Then
      Compras_XML!InformacoesAdicionais = Left(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/infAdic/infCpl").Item(0).Text, 100)
   End If
   
   Compras_XML!BaseICMS = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vBC").Item(0).Text) / 100)
   Compras_XML!ValorICMS = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vICMS").Item(0).Text) / 100)
   Compras_XML!BaseICMSST = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vBCST").Item(0).Text) / 100)
   Compras_XML!ValorICMSST = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vST").Item(0).Text) / 100)
   Compras_XML!ValorProdutos = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vProd").Item(0).Text) / 100)
   Compras_XML!ValorFrete = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vFrete").Item(0).Text) / 100)
   Compras_XML!ValorSeguro = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vSeg").Item(0).Text) / 100)
   Compras_XML!ValorDesconto = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vDesc").Item(0).Text) / 100)
   Compras_XML!ValorIsento = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vII").Item(0).Text) / 100)
   Compras_XML!ValorIPI = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vIPI").Item(0).Text) / 100)
   Compras_XML!ValorPIS = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vPIS").Item(0).Text) / 100)
   Compras_XML!ValorCOFINS = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vCOFINS").Item(0).Text) / 100)
   Compras_XML!ValorOutrasDespesas = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vOutro").Item(0).Text) / 100)
   Compras_XML!ValorNota = (CDbl(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vNF").Item(0).Text) / 100)
   
   Compras_XML!NumeroProtocolo = XMLDOC.getElementsByTagName("nfeProc/protNFe/infProt/nProt").Item(0).Text
   ''Compras_XML!datahoraprotocolo = XMLDOC.getElementsByTagName("nfeProc/protNFe/infProt/dhRecbto").Item(0).Text
   Compras_XML!enviada = 1
   'Compras_XML!ImportadoXML = True
   Compras_XML.Update
  '------------- FIM DO CADASTRO DA ENTRADA ------------------
      
      
      
       '-------- IMPORTANDO OS ITENS DA NOTA ---------
       'Contando quantos itens tem o nó det (detalhes)
       qtdProd = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/det").Length
      
       If qtdProd > 0 Then
          qtfracao = 1
          RsOpen Compras_Itens_XML, "SELECT TOP 1 * FROM EntradaEstoqueItens"
          RsOpen Produtos, "SELECT TOP 1 * FROM Produtos"
       End If

       For i = 0 To qtdProd - 1 'Varrendo todos os itens
        
           csticms = ""
           cstipi = ""
           cstpis = ""
           cstcofins = ""
           ICMSBase = 0
           ICMSAliquota = 0
           ICMSValor = 0
        
           csticms = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(0).childNodes(0).selectNodes("orig").Item(0).Text & XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(0).childNodes(0).selectNodes("CST").Item(0).Text
           
           If csticms = "" Then
              If XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/CRT").Item(0).Text <> "1" Then
                 csticms = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(1).childNodes(0).selectNodes("orig").Item(0).Text & XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(1).childNodes(0).selectNodes("CST").Item(0).Text
                 ICMSBase = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(1).childNodes(0).selectNodes("vBC").Item(0).Text, ".", ",", UM_A_UM))
                 ICMSAliquota = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(1).childNodes(0).selectNodes("pICMS").Item(0).Text, ".", ",", UM_A_UM))
                 ICMSValor = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(1).childNodes(0).selectNodes("vICMS").Item(0).Text, ".", ",", UM_A_UM))
              Else
                 csticms = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(1).childNodes(0).selectNodes("orig").Item(0).Text & XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(1).childNodes(0).selectNodes("CSOSN").Item(0).Text
                 If csticms = "" Then csticms = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(0).childNodes(0).selectNodes("orig").Item(0).Text & XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(0).childNodes(0).selectNodes("CSOSN").Item(0).Text
              End If
           Else
              If XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/CRT").Item(0).Text <> "1" Then
              'imposto/COFINS/COFINSNT/CST
                 ''ICMSBase = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(0).childNodes(0).selectNodes("vBC").Item(0).Text, ".", ",", UM_A_UM))
                 ICMSBase = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(0).childNodes(0).selectNodes("imposto/ICMS/ICMS00/vBC").Item(0).Text, ".", ",", UM_A_UM))
                 ICMSAliquota = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(0).childNodes(0).selectNodes("pICMS").Item(0).Text, ".", ",", UM_A_UM))
                 ICMSValor = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(0).childNodes(0).selectNodes("vICMS").Item(0).Text, ".", ",", UM_A_UM))
              End If
           End If
        
          'IMPOSTO IPI
           ''cstipi = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/IPI/IPITrib/CST").Item(0).Text

           ''If cstipi = "" Then
           ''   cstipi = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/IPI/IPINT/CST").Item(0).Text
           ''End If
        
            'IMPOSTO PIS
           ''cstpis = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/PIS/PISAliq/CST").Item(0).Text
           cstpis = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/PIS/PISNT/CST").Item(0).Text
           
           If cstpis = "" Then
              cstpis = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/PIS/PISNT/CST").Item(0).Text
              If cstpis = "" Then cstpis = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/PIS/PISOutr/CST").Item(0).Text
           End If
        
           'IMPOSTO COFINS
           ''cstcofins = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/COFINS/COFINSAliq/CST").Item(0).Text
           cstcofins = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/COFINS/COFINSNT/CST").Item(0).Text

           If cstcofins = "" Then
              cstcofins = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/COFINS/COFINSNT/CST").Item(0).Text
              If cstcofins = "" Then cstcofins = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/COFINS/COFINSOutr/CST").Item(0).Text
           End If
           
          cEAN = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cEAN").Item(0).Text
          
          
          If Not Vazio(cEAN) Then
          
             str_sql_item = "SELECT Codigo FROM Produtos WHERE Cod_Barra = '" & cEAN & "'"
             
             CodProduto = SQLExecutaRetorno(str_sql_item, "Codigo", 0)
             
             If CodProduto > 0 Then
                'str_sql_item = "SELECT Count(*) r FROM TbProdutosFornecedores WHERE IDProduto = " & codProduto & " AND Cod_prod_fornecedor = '" & XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text & "'"
                'If SQLExecutaRetorno(str_sql_item, "r", 0) = 0 Then
                'If Not PTab(Produtos_fornecedores, "FornecedorCodProdForn", Codigo_fornecedor, XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text) Then
                '    str_sql_item = "INSERT INTO [TbProdutosFornecedores] ([IDProduto], [IDFornecedor], [Cod_prod_fornecedor], [Fracionamento]) " & _
                '                   "     VALUES (" & CodProduto & ", " & Codigo_fornecedor & ", '" & XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text & "', 1)"
                '    SQLExecuta str_sql_item
                'Else
                '   qtfracao = Produtos_fornecedores!Fracionamento
                'End If
             Else
                str_sql_item = "Select IDProduto from ProdutosFornecedores where IDFornecedor = " & Codigo_fornecedor & " and Cod_prod_fornecedor = '" & XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text & "'"
                CodProduto = SQLExecutaRetorno(str_sql_item, "IDProduto", 0)
                
                If CodProduto > 0 Then
                   'PTab Produtos_fornecedores, "FornecedorCodProdForn", Codigo_fornecedor, XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text
                   qtfracao = Produtos_fornecedores!Fracionamento
                End If
             End If
          Else
             str_sql_item = "Select IDProduto from ProdutosFornecedores where IDFornecedor = " & Codigo_fornecedor & " and Cod_prod_fornecedor = '" & XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text & "'"
             CodProduto = SQLExecutaRetorno(str_sql_item, "IDProduto", 0)
             If CodProduto > 0 Then
                'PTab Produtos_fornecedores, "FornecedorCodProdForn", Codigo_fornecedor, XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text
                qtfracao = Produtos_fornecedores!Fracionamento
             End If
          End If
                
          'Monta String SQL para Pesquisa do Código do Produto na Tabela Produtos_Fornecedor
          ''str_sql_item = "Select * from ProdutosFornecedores where IDFornecedor = " & Codigo_fornecedor & " and Cod_prod_fornecedor = '" & XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text & "'"

          'Filta no BD
          ''RsOpen Produtos_fornecedores, str_sql_item
          
          'Verifica se o produto existe, caso não exista faz o cadastro
          ''If 1 = 2 Then
          ''   'str_sql_item = "SELECT * FROM Produtos WHERE [Código do produto] = '"
          ''   CodProduto = 1
          ''   Produtos.AddNew
          ''   Produtos!IdProduto = CodProduto
          ''   Produtos!NomeProduto = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/xProd").Item(0).Text
          ''   Produtos!NomeReduzido = Left$(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/xProd").Item(0).Text, 40)
          ''   Produtos!CodFabricante = Left$(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text, 20)
          ''   Produtos!UN = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/uCom").Item(0).Text
          ''   Produtos!EstoqueMin = 1
          ''   Produtos!ValorCusto = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vUnCom").Item(0).Text, ".", ",", UM_A_UM))
          ''   Produtos!IDFornecedorUlt = Codigo_fornecedor
          ''   Produtos!DtUltCompra = Date
          ''   'Produtos!Pro_cadastro = Date
          ''   Produtos!Aplicacao = "CRIADO AUTOMATICO VIA IMPORTAÇÃO XML"
          ''   Produtos!TipoTributo = 1
          ''   Produtos!Aliquota = ICMSAliquota
          ''   Produtos!CodNcm = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/NCM").Item(0).Text
          ''   Produtos!CodBarras = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cEAN").Item(0).Text
          ''   Produtos!ValorCompra = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vUnCom").Item(0).Text, ".", ",", UM_A_UM))
          ''   Produtos!pisCST = cstpis
          ''   Produtos!cofinsCST = cstcofins
          ''   Produtos!IPICST = cstipi
          ''   Produtos!CEST = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/CEST").Item(0).Text
          ''   Produtos.Update
             
             'cadastra produto no fornecedor
          ''   Produtos_fornecedores.AddNew
          ''   Produtos_fornecedores!IdProduto = CodProduto
          ''   Produtos_fornecedores!IDFornecedor = Codigo_fornecedor
          ''   Produtos_fornecedores!Cod_prod_fornecedor = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text
          ''   Produtos_fornecedores.Update
          ''Else
          ''   CodProduto = Produtos_fornecedores!IdProduto
          ''End If
    
           xCfop = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/CFOP").Item(0).Text
          ''If CodProduto > 0 Then
          ''   'PTab Produtos, "IDProduto", CodProduto
          ''    xCfop = VerificaCFOPEntrada(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/dest/enderDest/UF").Item(0).Text), UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").Item(0).Text), csticms, Produtos!CfopEnt)
          '' Else
          ''    xCfop = VerificaCFOPEntrada(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/dest/enderDest/UF").Item(0).Text), UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").Item(0).Text), csticms, "")
          ''End If
           'Compras_XML!CFOP = xCfop
           Compras_XML.Update
           
           'Lançando o Item no Banco
           Compras_Itens_XML.AddNew
           Compras_Itens_XML!CodigoNota = Compra
           'Compras_Itens_XML!Ite_documento = Compra
           Compras_Itens_XML!Item = SQLExecutaRetorno("SELECT ISNULL(MAX(Item),0) r FROM EntradaEstoqueItens WHERE CodigoNota = " & Compra, "r", 0) + 1
           Compras_Itens_XML!CodigoProduto = CodProduto
           Compras_Itens_XML!Referencia = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text
           Compras_Itens_XML!Genero = Codigo_fornecedor
           'Compras_itens_xml!Barras = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cEAN").Item(0).Text
           Compras_Itens_XML!NomeProduto = UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/xProd").Item(0).Text)
           Compras_Itens_XML!NCM = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/NCM").Item(0).Text
           'Compras_itens_xml!Ex_ipi = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/EXTIPI").Item(0).Text
           Compras_Itens_XML!CFOP = xCfop
           Compras_Itens_XML!UnidadeComercial = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/uCom").Item(0).Text
           Compras_Itens_XML!QuantidadeComercial = (CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/qCom").Item(0).Text, ".", ",", UM_A_UM)) * qtfracao)
           Compras_Itens_XML!ValorUnitarioComercializacao = (CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vUnCom").Item(0).Text, ".", ",", UM_A_UM)) / qtfracao)
           Compras_Itens_XML!ValorTotalBruto = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vProd").Item(0).Text, ".", ",", UM_A_UM))
           'Compras_Itens_XML!Ite_frete = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vFrete").Item(0).Text, ".", ",", UM_A_UM))
           'Compras_itens_xml!Vlr_seguro = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vSeg").Item(0).Text
           ''Compras_Itens_XML!ValorDesconto = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vDesc").Item(0).Text, ".", ",", UM_A_UM))
           'Compras_Itens_XML!Ite_outimp = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vOutro").Item(0).Text, ".", ",", UM_A_UM))
           'Compras_itens_xml!Origem = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(0).childNodes(0).selectNodes("orig").Item(0).Text
           Compras_Itens_XML!CST = csticms
           Compras_Itens_XML!vBC = ICMSBase
           Compras_Itens_XML!pICMS = ICMSAliquota
           Compras_Itens_XML!vICMS = ICMSValor
           Compras_Itens_XML!IPICST = cstipi
           ''Compras_Itens_XML!IPIpIPI = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/IPI/IPITrib/pIPI").Item(0).Text, ".", ",", UM_A_UM))
           'Compras_itens_xml!Vlr_base_ipi = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/IPI/IPITrib/vBC").Item(0).Text
           'Compras_itens_xml!Valor_IPI = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/IPI/IPITrib/vIPI").Item(0).Text
           Compras_Itens_XML!pisCST = cstpis
           'Compras_itens_xml!Per_ent_pis = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/PIS/PISAliq/pPIS").Item(0).Text
           'Compras_itens_xml!Vlr_base_pis = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/PIS/PISAliq/vBC").Item(0).Text
           'Compras_itens_xml!Vlr_pis = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/PIS/PISAliq/vPIS").Item(0).Text
           Compras_Itens_XML!cofinsCST = cstcofins
           'Compras_itens_xml!Per_ent_cofins = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/COFINS/COFINSAliq/pCOFINS").Item(0).Text
           'Compras_itens_xml!Vlr_base_cofins = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/COFINS/COFINSAliq/vBC").Item(0).Text
           'Compras_itens_xml!Vlr_cofins = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/COFINS/COFINSAliq/vCOFINS").Item(0).Text
           'Compras_Itens_XML!Com_data = Date
           'Compras_Itens_XML!Ite_hora = Time
           'Compras_Itens_XML!ImportadoXML = True
           'Compras_Itens_XML!ValorVendaNovo = SQLExecutaRetorno("SELECT ValorVenda FROM Produtos WHERE IDProduto = " & CodProduto, "ValorVenda", 0)
           Compras_Itens_XML!EAN = cEAN
           If CodProduto = 0 Then
              'Compras_Itens_XML!IDGrupo = 1
              'Compras_Itens_XML!IdLinhaProd = 1
              'Compras_Itens_XML!IdFabricante = 1
              'Compras_Itens_XML!ICMSCSTSaida = csticms
              'Compras_Itens_XML!Aliq_IcmsSaida = ICMSAliquota
              'Compras_Itens_XML!PISCSTSaida = cstpis
              'Compras_Itens_XML!COFINSCSTSaida = cstcofins
              'Compras_Itens_XML!IPICSTSaida = "53"
           End If
           Compras_Itens_XML.Update
   
    Next i
    '------------- FIM DA IMPORTAÇÃO DOS ITENS ------------------
    

    '------------- IMPORTANDO AS DUPLICATAS ---------
    
    i = 0
    qtdDup = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/cobr/dup").Length 'Contando quantas parcelas existem
    
    GoTo pulaParcelas
    
    If xContasPagar Then GoTo LancarParcelas
    
    If qtdDup > 0 Then
       
       If MsgBox("LANÇAR AS PARCELAS NO CONTAS A PAGAR?", vbQuestion + vbYesNo, vgAtencao) = vbYes Then
LancarParcelas:
          Compras_XML!ParcelasXML = True
          Compras_XML.Update
          For i = 0 To qtdDup - 1 'Varrendo todas parcelas
               str_sql = "EXEC tbContasPagarIncluir " & Compra & ",0,'" & CStr(i + 1) + "/" + CStr(qtdDup) & "', " & FdtSQL(Date) & ", " & Codigo_fornecedor & ", '" & _
                                                 "REF. COMPRA " & LPad(Compra, 6, "0") & " - DUP: " & XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").Item(i).selectNodes("nDup").Item(0).Text & "', " & _
                                                 FSQL(CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").Item(i).selectNodes("vDup").Item(0).Text, ".", ",", UM_A_UM))) & ", " & _
                                                 FdtSQL(XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").Item(i).selectNodes("dVenc").Item(0).Text) & ", 0, '" & vgPWUsuario & "', '" & _
                                                 NomeEstacao & "', 0, 'BL', '', '', 0, '', 0, " & vgFilial & ", 0, 0, 0, '" & cpConta.Value & "', 0, '', " & vgFilial & ", '" & Format(XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").Item(i).selectNodes("dVenc").Item(0).Text, "yyyy/mm") & "', 0, " & _
                                                 "'" & XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").Item(i).selectNodes("nDup").Item(0).Text & "'"
               msgResultado = SQLExecuta(str_sql)
               If Not Vazio(msgResultado) Then
                  MsgBox msgResultado, vbCritical
                  Exit Sub
               End If
          Next i
          str_sql = "UPDATE TbContasPagar SET excluido = 1 WHERE IdEntradaEst = " & Compra
          msgResultado = SQLExecuta(str_sql)
          If Not Vazio(msgResultado) Then
             MsgBox msgResultado, vbCritical
             Exit Sub
          End If
       End If
    '------------- FIM DA IMPORTAÇÃO DAS DUPLICATAS ------------------
    End If

pulaParcelas:
   MsgBox "NF-e importada com sucesso!", vbInformation + vbOKOnly, "Importa XML"

   str_sql = "SELECT Count(*) r FROM EntradaEstoqueItens WHERE CodigoNota = " & Compra & " AND CodigoProduto = 0"
   If SQLExecutaRetorno(str_sql, "r", 0) > 0 Then
       xCaminhoXML = arqXML
   End If
   
   RsOpen TbEntrada, "SELECT *, " & _
                     "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                     "FROM EntradaEstoque WHERE CodigoNota = " & Compra
   Load_Controls
   Frm_NF.Tab = 0
   
End Sub

Private Function Inserir_Fornecedor() As Boolean
Dim sSQL As String

Dim varCodFornecedor As Integer
Dim varRAZAOForn As String
Dim varFantasiaFor As String
Dim varCNPJforn As String
Dim varIEforn As String
Dim varENDERECOforn As String
Dim varNUMforn As Integer
Dim varBairroforn As String
Dim varCidadeforn As String
Dim varEstadoforn As String
Dim varCEPform As String
Dim varCODIBGEform As Integer

varCodFornecedor = SQLExecutaRetorno("SELECT ISNULL(MAX(Codigo),0) r FROM fornecedor", "r", 0) + 1
varRAZAOForn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xNome").Item(0).Text), 35)
varFantasiaFor = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xFant").Item(0).Text), 35)
If tipoEmp = 0 Then varCNPJforn = CPF
If tipoEmp = 1 Then varCNPJforn = CNPJ
varIEforn = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/IE").Item(0).Text
varENDERECOforn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xLgr").Item(0).Text), 40) & ", " & Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").Item(0).Text), 8)
varBairroforn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xBairro").Item(0).Text), 20)
varCidadeforn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xMun").Item(0).Text), 20)
varEstadoforn = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").Item(0).Text)
varCEPform = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/CEP").Item(0).Text
varNUMforn = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").Item(0).Text
varCODIBGEform = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/cMun").Item(0).Text


'Comando de inclusão
sSQL = "INSERT INTO fornecedor (" & _
   "codigo, " & _
   "RAZAO, " & _
   "Fantasia, " & _
   "CNPJ, " & _
   "IE, " & _
   "ENDERECO, " & _
   "bairro, " & _
   "Cidade, " & _
   "Estado, " & _
   "CEP, " & _
   "Numero, " & _
   "CodigoIBGE) VALUES ("

sSQL = sSQL & _
   varCodFornecedor & ", '" & varRAZAOForn & "', '" & varFantasiaFor & "', '" & varCNPJforn & "', '" & _
   varIEforn & "', '" & varENDERECOforn & "', '" & varBairroforn & "', '" & varCidadeforn & "', '" & _
   cboEstado.Text & "', '" & mskTelefone1.Text & "', '" & mskTelefone2.Text & "', '" & mskFax.Text & "', '" & _
   varEstadoforn & "', '" & varCEPform & "', '" & varNUMforn & "', '" & varCODIBGEform & "')"

dbData.Execute (sSQL)

   Fornecedores.AddNew
   Codigo_fornecedor = SQLExecutaRetorno("SELECT ISNULL(MAX(Codigo),0) r FROM fornecedor", "r", 0) + 1
   
   Fornecedores!Codigo = Codigo_fornecedor
   Fornecedores!Razao = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xNome").Item(0).Text), 35)
   Fornecedores!Fantasia = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xFant").Item(0).Text), 35)
   If tipoEmp = 0 Then Fornecedores!CNPJ = CPF
   If tipoEmp = 1 Then Fornecedores!CNPJ = CNPJ
   Fornecedores!IE = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/IE").Item(0).Text
   Fornecedores!Endereco = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xLgr").Item(0).Text), 40) & ", " & Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").Item(0).Text), 8)
   Fornecedores!bairro = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xBairro").Item(0).Text), 20)
   Fornecedores!Cidade = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xMun").Item(0).Text), 20)
   Fornecedores!Estado = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").Item(0).Text)
   Fornecedores!CEP = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/CEP").Item(0).Text
   Fornecedores.Update

End Function

Private Sub Combo1_Change()

End Sub

Private Sub Form_Activate()
If lblProdutoCadastrado.Caption = "Produto Inexistente" Then
    txtCodProdExist = ""
    txtCodBarraExist = ""
    txtDescricaoExist = ""
    txtUnidMedExist = ""
    txtQuantExist = ""
    txtValorExist = ""
    txtNCMExist = ""
    txtCSTExist = ""
    txtCFOPExist = ""
    txtCodBarraSel = ""
    txtDescricaoSel = ""
    txtUnidMedSel = ""
    txtQuantSel = ""
    txtValorSel = ""
    txtNCMSel = ""
    txtCSTSel = ""
    txtCFOPSel = ""
    txtQuantBruto = ""
    txtValorBruto = ""
    txtItem = ""
    lblProdutoCadastrado.Caption = ""
    cmdAtualEAN.Enabled = False
    cmdAtualDesc.Enabled = False
    cmdAtualUnid.Enabled = False
    cmdAtualNCM.Enabled = False
    cmdAtualProd.Enabled = False
End If
End Sub

Private Sub GridEntradaItens_Click()

'RsOpen TbEntradaItens, "SELECT EAN, NOMEPRODUTO, item, UnidadeComercial, QuantidadeComercial, ValorUnitarioComercializacao, NCM, CST, CFOP " & _
                "FROM EntradaEstoqueItens WHERE  Item = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 18) & " And CodigoNota = " & txtCodNota.Text

           ' If .TextMatrix(i, 19) = "SIM" Then
               '.CellForeColor = vbRed
            'Else
               '.CellForeColor = vbBlack
            'End If
End Sub

Private Sub GridNotasItens_DblClick()
If txtCodNota.Text = "" Then Exit Sub

If GridNotasItens.TextMatrix(GridNotasItens.Row, 19) = "SIM" Then
    MsgBox "Produto já adicionado no estoque fisico!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

RsOpen TbEntradaItens, "SELECT EAN, NOMEPRODUTO, item, UnidadeComercial, QuantidadeComercial, ValorUnitarioComercializacao, NCM, CST, CFOP " & _
                "FROM EntradaEstoqueItens WHERE  Item = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 18) & " And CodigoNota = " & txtCodNota.Text

If Len(TbEntradaItens("EAN")) > 13 Then
    txtCodBarraSel = Format(Val(TbEntradaItens("EAN")), "@")
Else
    txtCodBarraSel = Format(TbEntradaItens("EAN"), "@")
End If

txtDescricaoSel = Format(TbEntradaItens("NOMEPRODUTO"), "@")
txtUnidMedSel = Format(TbEntradaItens("UnidadeComercial"), "@")
txtQuantSel = Format(TbEntradaItens("QuantidadeComercial"), "@")
txtValorSel = Format(TbEntradaItens("ValorUnitarioComercializacao"), "##,##0.00")
txtNCMSel = Format(TbEntradaItens("NCM"), "@")
txtCSTSel = Format(TbEntradaItens("CST"), "@")
txtCFOPSel = Format(TbEntradaItens("CFOP"), "@")
txtQuantBruto = Format(TbEntradaItens("QuantidadeComercial"), "@")
txtValorBruto = Format(TbEntradaItens("ValorUnitarioComercializacao"), "##,##0.00")
txtItem = Format(TbEntradaItens("item"), "@")

LimparObjetosProdutoSel
Call MostrarProdutoExistente
End Sub


Private Sub mskConNotaFinal_GotFocus()
SelectControl mskConNotaFinal
End Sub

Private Sub mskConNotaFinal_KeyPress(KeyAscii As Integer)
mskConNotaFinal.Mask = "##/##/##"
End Sub


Private Sub mskConNotaFinal_LostFocus()
If mskConNotaFinal.Text = "" Or mskConNotaFinal.Text = "__/__/__" Then
   mskConNotaFinal.Mask = ""
   mskConNotaFinal.Text = ""
   Exit Sub
Else
   If IsDate(mskConNotaFinal.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskConNotaFinal.SetFocus
      SelectControl mskConNotaFinal
   End If
End If
End Sub


Private Sub mskConNotaInicial_GotFocus()
SelectControl mskConNotaInicial
End Sub


Private Sub mskConNotaInicial_KeyPress(KeyAscii As Integer)
mskConNotaInicial.Mask = "##/##/##"
End Sub


Private Sub mskConNotaInicial_LostFocus()
If mskConNotaInicial.Text = "" Or mskConNotaInicial.Text = "__/__/__" Then
   mskConNotaInicial.Mask = ""
   mskConNotaInicial.Text = ""
   Exit Sub
Else
   If IsDate(mskConNotaInicial.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskConNotaInicial.SetFocus
      SelectControl mskConNotaInicial
   End If
End If

End Sub

Private Sub mskEmissao_KeyPress(KeyAscii As Integer)
mskEmissao.Mask = "##/##/####"
End Sub



Private Sub optCXUnid_Click()
cmdUsarUnidMedida.Visible = False
cmdFracionar.Visible = True
txtQuantUnid.Enabled = True
lblFracao.Enabled = True
End Sub

Private Sub optPesoMetro_Click()
cmdUsarUnidMedida.Visible = False
cmdFracionar.Visible = True
txtQuantUnid.Enabled = True
lblFracao.Enabled = True
End Sub


Private Sub optUnidUnid_Click()
cmdUsarUnidMedida.Visible = True
cmdFracionar.Visible = False
txtQuantUnid.Enabled = False
lblFracao.Enabled = False
End Sub






















Private Sub txtCodEntrada_Change()
Mostrar_Itens
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then SendK vbKeyTab
If KeyCode = vbKeyUp Then
    SendK vbKeyTab
    KeyCode = 0
ElseIf KeyCode = vbKeyDown Then
    SendK vbKeyTab
    KeyCode = 0
End If

If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
Set moCombo = New cComboHelper
Frm_NF.Tab = 0

cboFiltroNota_GotFocus
cboFiltroNota.ListIndex = 5

cboOrdem_GotFocus
cboOrdem.ListIndex = 0

cboConNotaMes.Text = Format(Date, "mmmm")
cboConNotaAno.Text = Year(Date)
cmdExibirConNotas_Click

'MostrarNotasInicias
cmdNovo.Enabled = True
cmdUsarUnidMedida.Visible = True
cmdFracionar.Visible = False
cmdImportarXML.Enabled = True
txtCodUsuario.Text = Tela_Principal.txtCodFuncionario.Text
End Sub

Private Sub GridNotas_DblClick()
RsOpen TbEntrada, "SELECT *, " & _
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                "FROM EntradaEstoque WHERE CodigoNota = " & GridNotas.TextMatrix(GridNotas.Row, 1)
Load_Controls
Frm_NF.Tab = 0
Tab_Entrada.Tab = 0
End Sub





Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
Dim i As Integer, Valor As Currency

Valor = 0
For i = 0 To var_Grid.rows - 1
   If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
      Valor = Valor + CCur(var_Grid.TextMatrix(i, Col))
   End If
Next

SomaGrid = Valor
End Function









Private Sub Mostrar_Pedido(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then

Dim totalRegistros As Long

    'buscar Numero e codigo da nota (autopreenchimento)
    RsOpen TbNotaPedido, "SELECT *, " & _
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                "FROM EntradaEstoque"
                
    If TbNotaPedido.RecordCount > 0 Then totalRegistros = TbNotaPedido.RecordCount
        
    'Clear_Controls
    
    'INICIO DO PREENCHIMENTOS DOS OBJETOS
    If TbNotaPedido.EOF And TbNotaPedido.BOF Then
        txtNumNota.Text = "1"
        txtCodNota.Text = "1"
    Else
        TbNotaPedido.MoveLast
        txtNumNota.Text = TbNotaPedido("NumeroNota") + 1
        txtCodNota.Text = TbNotaPedido("CodigoNota") + 1
    End If

    
'    txtNumNota = Format(rTabela("NumeroNota"), "@")
'    txtCodNota = Format(rTabela("CodigoNota"), "@")
    TxtCodFornecedor = Format(rTabela("COD_CLIENTE"), "@")
    cboFornecedor = Format(rTabela("varnome"), "@")
    'cboDestOperacao = Format(5102, "@")
    cboDestOperacao.Text = "1 - Operação Interna"
    'Text1 = Format(rTabela("InformacoesAdicionais"), "@")
    cboNatureza = "VENDA DE MERCADORIA"
    'mskEmissao = Format(Date, "dd/mm/yyyy")
    'mskSaida = Format(Date, "dd/mm/yyyy")
    'mskHora = Format(Time(), "HH:MM:ss")
    
    'mostrar a parte dos transporte
    'txtCodTransporte = Format(0, "@")
    'cboTransporte = Format(0, "@")   'ver 2
    'txtPlaca = Format(0, "@")
    'Text10 = Format(0, "@")
    'Text11 = Format(0, "@")
    'Text12 = Format(0, "@")
    'txtVolNumeracao = Format(0, "@")
    'Text15 = Format(0, "@")
    
    'Text19 = Format(0, "##,##0.00")
    'Text16 = Format(0, "##,##0.00")
    txtValorFrete = Format(0, "##,##0.00")
    Text22 = Format(0, "##,##0.00")
    Text21 = Format(0, "##,##0.00")
    'Text17 = Format(0, "@")
    'Text18 = Format(0, "@")
    'Text23 = Format(0, "@")
    cboModFrete = "9 - SEM FRETE"
    
    Text30 = Format(0, "@")
    Text31 = Format(0, "@")
    Text32 = Format(0, "@")
    cboIndicadorPagamento.Text = "0 - Pagamento à vista"
    cboFormatoDANFe.Text = "1 - Retrato"
    cboTipoEmissao.Text = "1 - Normal"
End If
End Sub

Private Sub Calcular_Desconto()
Dim varTotalUnid As Currency
Dim varTotalSemDesc As Currency
Dim varTotalComdesc As Currency
Dim varQuant As Double
Dim varDesc As Currency

If txtDesc.Text = "" Then txtDesc.Text = "0"
If txtQuant.Text = "" Then txtQuant.Text = "1": SelectControl txtDesc
If txtValor.Text = "" Then txtValor.Text = "0"

varTotalUnid = txtValor.Text

varDesc = txtDesc.Text
varQuant = txtQuant.Text
varTotalSemDesc = CCur(varTotalUnid) * CDbl(varQuant)
varTotalComdesc = CCur(varTotalUnid) * CDbl(varQuant) - CCur(varDesc)

If txtValor.Text = "0" Or txtValor.Text = "0,00" Then Exit Sub
If txtDesc.Text = "" Then txtDesc.Text = Format(0, ocMONEY)

If txtDesc.Text <> "0,00" Then
    txtSubtotal.Text = Format(varTotalComdesc, ocMONEY)
Else
    txtSubtotal.Text = Format(varTotalSemDesc, ocMONEY)
End If
End Sub


Private Sub txtDescricaoExist_Change()
If txtDescricaoExist.Text = "" Then
    frmCusto2.Enabled = False
    frmVenda.Enabled = False
Else
    frmCusto2.Enabled = True
    frmVenda.Enabled = True
End If
End Sub


Private Sub txtMargemAP_Validate(Cancel As Boolean)
txtMargemAP_LostFocus
End Sub

Private Sub txtMargemAV_Validate(Cancel As Boolean)
txtMargemAV_LostFocus
End Sub

Private Sub txtMargemVP_Validate(Cancel As Boolean)
txtMargemVP_LostFocus
End Sub

Private Sub txtnumnota_GotFocus()
txtNumNota.SelStart = 0
txtNumNota.SelLength = Len(txtNumNota)
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

If Right(txtMargemAP.Text, 1) = "%" Then
   varMargemAP = Left$(txtMargemAP.Text, Len(txtMargemAP.Text) - 1)
Else
    varMargemAP = txtMargemAP.Text
End If


txtMargemAP.Text = FormatNumber(varMargemAP, 2) & "%"

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

If Right(txtMargemAV.Text, 1) = "%" Then
   varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)
Else
    varMargemAV = txtMargemAV.Text
End If

txtMargemAV.Text = FormatNumber(varMargemAV, 2) & "%"

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

If Right(txtMargemVP.Text, 1) = "%" Then
   varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)
Else
    varMargemVP = txtMargemVP.Text
End If

txtMargemVP.Text = FormatNumber(varMargemVP, 2) & "%"

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

txtMargemVV.Text = FormatNumber(varMargemVV, 2) & "%"
If txtMargemVP.Text = "" Then txtMargemVP.Text = txtMargemVV.Text
If txtMargemAV.Text = "" Then txtMargemAV.Text = txtMargemVV.Text
If txtMargemAP.Text = "" Then txtMargemAP.Text = txtMargemVV.Text
CalcularPrecos
lblAviso.Visible = False
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
Dim C As Currency

a = txtCusto.Text
B = txtValorAP.Text
C = ((B - a) / a) * 100

txtMargemAP.Text = FormatNumber(C, 2) & "%"
txtValorAP.Text = Format(txtValorAP.Text, ocMONEY)
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
Dim C As Currency

a = txtCusto.Text
B = txtValorAV.Text
C = ((B - a) / a) * 100

txtMargemAV.Text = FormatNumber(C, 2) & "%"
txtValorAV.Text = Format(txtValorAV.Text, ocMONEY)
End Sub


Private Sub txtnumnota_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub


Private Sub mskEmissao_GotFocus()
If mskEmissao.Text = "" Then mskEmissao = Format(Date, "dd/mm/yyyy")

mskEmissao.SelStart = 0
mskEmissao.SelLength = Len(mskEmissao)
End Sub
Private Sub txtCodNota_Change()
Mostrar_ItensNota
End Sub




Private Function MostraStatus_F9() As String
 On Error Resume Next
 If TbEntrada("Denegada") Then
   MostraStatus.ForeColor = vbRed
   MostraStatus_F9$ = "Denegada"  'Deve retornar uma expressão caractere
 ElseIf TbEntrada("Enviada") And Not TbEntrada("Cancelada") Then
   MostraStatus.ForeColor = vbBlue
   MostraStatus_F9$ = "Transmitida/Autorizada"  'Deve retornar uma expressão caractere
 ElseIf TbEntrada("Cancelada") Then
   MostraStatus.ForeColor = vbRed
   MostraStatus_F9$ = "Cancelada"  'Deve retornar uma expressão caractere
 Else
   MostraStatus.ForeColor = vbBlack
   MostraStatus_F9$ = "Em Digitação"  'Deve retornar uma expressão caractere
 End If
End Function

Private Sub LimparGridNotas()
Dim i As Integer, j As Integer

With GridNotas
   .Visible = False
   .Redraw = False
   
   .Clear
   .Cols = 9
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 1000
   .ColWidth(3) = 1100
   .ColWidth(4) = 4700
   .ColWidth(5) = 1100
   .ColWidth(6) = 1350
   .ColWidth(7) = 1350
   .ColWidth(8) = 1350
   
   'CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota
   .TextMatrix(0, 1) = "CÓD"
   .TextMatrix(0, 2) = "NUM NF"
   .TextMatrix(0, 3) = "EMISSÃO"
   .TextMatrix(0, 4) = "FORNECEDOR"
   .TextMatrix(0, 5) = "VALOR"
   .TextMatrix(0, 6) = "STATUS"
   .TextMatrix(0, 7) = "SITUAÇÃO"
   .TextMatrix(0, 8) = "PENDÊNCIAS"
   
   'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   'centralizar o titulo
   For i = 0 To .Cols - 1
      .Row = 0
      .Col = i
      .CellAlignment = flexAlignCenterCenter
   Next
   
   i = 1
   
   'ALINHAMENTO
   .ColAlignment(0) = 1
   .ColAlignment(1) = 1
   .ColAlignment(2) = 1
   .ColAlignment(3) = 1
   .ColAlignment(4) = 1
   .ColAlignment(5) = 6
   .ColAlignment(6) = 2
   .ColAlignment(7) = 1
   .ColAlignment(8) = 1
   
   .rows = .rows - 1
   
   'MUDAR COR DE FONTE DA COLUNA
   For i = 1 To .rows - 1
      .Row = i
      .Col = 2
      .CellFontBold = True
   Next
   
   'MUDAR COR DE FONTE DA COLUNA
   For i = 1 To .rows - 1
      .Row = i
      .Col = 3
      .CellForeColor = &HC0&
      .CellFontBold = True
   Next
           
  'GridNotas.ColWidth(0) = 400
   'GridNotas.Rows = 11
   GridNotas.Col = 0
         
   .Visible = True
   .Redraw = True
End With
End Sub

Private Sub LimparGridItensNota()
   Dim i As Integer
   
   With GridNotasItens
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 8
      .rows = 2
      
      .ColWidth(0) = 200
      .ColWidth(1) = 2000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1200
      .ColWidth(6) = 1500
      .ColWidth(7) = 1500
      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      .TextMatrix(0, 1) = "CÓDIGO"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "CST"
      .TextMatrix(0, 4) = "UND"
      .TextMatrix(0, 5) = "QTDE"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "SUBTOTAL"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 2
      .ColAlignment(6) = 2
      .ColAlignment(7) = 2
      .rows = .rows + 1
      
      i = i + 1
      .rows = .rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 2
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'GridNotasItens.ColWidth(0) = 400
      'GridNotasItens.Rows = 11
      'GridNotasItens.Col = 0
      
      .Visible = True
      .Redraw = True
   End With
End Sub


Public Sub FormatarGridNotas(rTabela As ADODB.Recordset)
Dim i As Integer, j As Integer

With GridNotas
   .Visible = False
   .Redraw = False
   
   .Clear
   .Cols = 9
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 1000
   .ColWidth(3) = 1100
   .ColWidth(4) = 4700
   .ColWidth(5) = 1100
   .ColWidth(6) = 1350
   .ColWidth(7) = 1350
   .ColWidth(8) = 1350
   
   'CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota
   .TextMatrix(0, 1) = "CÓD"
   .TextMatrix(0, 2) = "NUM NF"
   .TextMatrix(0, 3) = "EMISSÃO"
   .TextMatrix(0, 4) = "FORNECEDOR"
   .TextMatrix(0, 5) = "VALOR"
   .TextMatrix(0, 6) = "STATUS"
   .TextMatrix(0, 7) = "SITUAÇÃO"
   .TextMatrix(0, 8) = "PENDÊNCIAS"
   
   'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   'centralizar o titulo
   For i = 0 To .Cols - 1
      .Row = 0
      .Col = i
      .CellAlignment = flexAlignCenterCenter
   Next
   
   i = 1
   
   'ALINHAMENTO
   .ColAlignment(0) = 1
   .ColAlignment(1) = 6
   .ColAlignment(2) = 1
   .ColAlignment(3) = 4
   .ColAlignment(4) = 1
   .ColAlignment(5) = 6
   .ColAlignment(6) = 4
   .ColAlignment(7) = 4
   .ColAlignment(8) = 6
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("CodigoNota")
         .TextMatrix(.rows - 1, 2) = Format(rTabela("NumeroNota"), "000000")
         .TextMatrix(.rows - 1, 3) = Format(rTabela("DataEmissao"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 4) = rTabela("RazaoSocial")
         .TextMatrix(.rows - 1, 5) = Format(rTabela("ValorNota"), ocMONEY)
         .TextMatrix(.rows - 1, 6) = rTabela("Status")
         .TextMatrix(.rows - 1, 7) = rTabela("SITUACAO")
         .TextMatrix(.rows - 1, 8) = rTabela("VSOMAPENDENCIA")
         rTabela.MoveNext
         .rows = .rows + 1
         i = i + 1
      Loop
   End If
   
   .rows = .rows - 1
   
   'MUDAR COR DE FONTE DA COLUNA
   For i = 1 To .rows - 1
      .Row = i
      .Col = 2
      .CellFontBold = True
   Next
   
   'MUDAR COR DE FONTE DA COLUNA
   For i = 1 To .rows - 1
      .Row = i
      .Col = 3
      .CellForeColor = &HC0&
      .CellFontBold = True
   Next
           
  'GridNotas.ColWidth(0) = 400
   'GridNotas.Rows = 11
   GridNotas.Col = 0
         
   .Visible = True
   .Redraw = True
End With

lblSomaNota.Caption = Format(SomaGrid(GridNotas, 5), ocMONEY)
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
Dim C As Currency

a = txtCusto.Text
B = txtValorVP.Text
C = ((B - a) / a) * 100

txtMargemVP.Text = FormatNumber(C, 2) & "%"
txtValorVP.Text = Format(txtValorVP.Text, ocMONEY)
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
Dim C As Currency

a = txtCusto.Text
B = txtValorVV.Text
C = ((B - a) / a) * 100

txtMargemVV.Text = FormatNumber(C, 2) & "%"
txtValorVV.Text = Format(txtValorVV.Text, ocMONEY)
End Sub


