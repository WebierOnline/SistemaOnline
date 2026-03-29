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
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(4)=   "Label19"
      Tab(1).Control(5)=   "Label20"
      Tab(1).Control(6)=   "Label21"
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(9)=   "Label22"
      Tab(1).Control(10)=   "Label18"
      Tab(1).Control(11)=   "Label17"
      Tab(1).Control(12)=   "Label10"
      Tab(1).Control(13)=   "Label12"
      Tab(1).Control(14)=   "Label13"
      Tab(1).Control(15)=   "Label9"
      Tab(1).Control(16)=   "lblTotalNota"
      Tab(1).Control(17)=   "lblSomaNota"
      Tab(1).Control(18)=   "GridNotas"
      Tab(1).Control(19)=   "Frame3"
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
         Tab(0).Control(10)=   "cmdVinculos"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).ControlCount=   11
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
         Begin VB.CommandButton cmdVinculos 
            Caption         =   "Ver &Vinculos"
            Height          =   375
            Left            =   12960
            TabIndex        =   230
            Top             =   2760
            Width           =   2175
         End
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
               Caption         =   "Varejo - à vista"
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
               Caption         =   "Varejo - à Prazo"
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
               Caption         =   "Atac. - à Vista"
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
               Caption         =   "Atac. - à Prazo"
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
            Caption         =   "Atacado - à Prazo"
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
            Caption         =   "Atacado - à Vista"
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
            Caption         =   "Varejo - à Prazo"
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
            Caption         =   "Varejo - à vista"
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
            TextSave        =   "11:46"
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
    '    cboDestOperacao.Text = "1 - Operaaao Interna"
    'ElseIf TbEntrada("CodigoNatureza") = 2 Then
    '    cboDestOperacao.Text = "2 - Operaaao Interestadual"
    'ElseIf TbEntrada("CodigoNatureza") = 3 Then
    '    cboDestOperacao.Text = "3 - Operaaao com Exterior"
    'End If
    
    'If TbEntrada("IndicadorIEDestinatario") = 1 Then
    '    cboTipoContribuinte.Text = "1 - CONTRIBUINTE ICMS"
    'ElseIf TbEntrada("IndicadorIEDestinatario") = 2 Then
    '    cboTipoContribuinte.Text = "2 - CONTRIBUINTE ISENTO"
    'ElseIf TbEntrada("IndicadorIEDestinatario") = 9 Then
    '    cboTipoContribuinte.Text = "9 - NaO CONTRIBUINTE"
    'End If

    'If TbEntrada("ConsumidorFinal") = False Then
    '    cboConsumidorFinal.Text = "0 - NaO"
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

    'habilita/desabilita botaes
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

sSQL = "SELECT *, (CASE WHEN adicionada = 1 THEN 'SIM' ELSE 'NaO' END) AS VarAdicionada FROM EntradaEstoqueItens WHERE CodigoNota = " & Val(txtCodNota.Text)
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
      
      .TextMatrix(0, 1) = "CaD."
      .TextMatrix(0, 2) = "DESCRIaaO"
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

      
      'colocar os cabeaalho em negrito
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
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
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
        lblProdutoCadastrado.Caption = "Esse produto nao foi encontrado no seu banco de dados!"
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
        If ShowMsg("O Produto: " & txtDescricaoSel.Text & " nao foi encontrado! Voca deseja localizar ele manualmente?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then cmdLocalizarManual_Click
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
        '    MsgBox "produto nao fracionado1", vbInformation, "Aviso do Sistema"
        '    'frmProdFracionado.Visible = False
        'End If
'End If
        Exit Sub
    End If
End If

If TbEntradaItens("EAN") = "SEM GTIN" Then
        lblProdutoCadastrado.Caption = "Esse produto nao foi encontrado no seu banco de dados!"
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
        If ShowMsg("O Produto: " & txtDescricaoSel.Text & " nao foi encontrado! Voca deseja localizar ele manualmente?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then cmdLocalizarManual_Click

Else
    RsOpen TbProdutos, "SELECT EAN, DESCRICAO, UNID_MEDIDA, QUANT_ESTOQUE, NCM, ICMSCST, CFOP, CODIGO, FRACIONADO, QUANT_FRACAO, CODPROD_FRACAO, ATIVO,  " & _
                       "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
                       "FROM produtos WHERE COD_BARRA = '" & txtCodBarraSel.Text & "' AND ATIVO = 1"
    
    
    If TbProdutos.EOF And TbProdutos.BOF Then
        lblProdutoCadastrado.Caption = "Esse produto nao foi encontrado no seu banco de dados!"
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
        If ShowMsg("O Produto: " & txtDescricaoSel.Text & " nao foi encontrado! Voca deseja localizar ele manualmente?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then cmdLocalizarManual_Click
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
        '    MsgBox "produto nao fracionado2", vbInformation, "Aviso do Sistema"
        '    frmProdFracionado.Visible = False
        'End If
    End If
End If
End Sub

Private Sub PreencherGridNotas()
Dim totalRegistros As Long

On Error GoTo ErrLoad

RsOpen TbConsulta, "SELECT *, " & _
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status " & _
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

Public Sub CriarEntradaEstoque(ByVal lCompra As Long)
Dim sSQL As String
Dim sData As String
Dim sHora As String
Dim msgResultado As String

sData = Format$(Date, ocDATA)
sHora = Format$(Time(), ocHORA)

Dim VarCodEntradaNovo As Long
VarCodEntradaNovo = SQLExecutaRetorno("SELECT ISNULL(MAX(Codigo),0) r FROM produtos_entrada", "r", 0) + 1

'-- colunas (parte 1: basicas)
sSQL = "INSERT INTO produtos_entrada (" & _
       "CODIGO, DATA_ENTRADA, HORA_ENTRADA, NOTAFISCAL, VALOR, " & _
       "DataEmissao, DataSaida, HoraSaida, " & _
       "COD_FORNECEDOR, COD_TRANSPORTADORA, VALOR_FRETE, Chave, XML, TIPO_FRETE, " & _
       "CodigoNota, SerieNF, ModeloNF, NaturezaOperacao, RazaoSocial, CNPJ_CPF, " & _
       "ValorProdutos, ValorDesconto, ValorIPI, ValorICMS, ValorICMSST) "
'-- SELECT valores (parte 2)
sSQL = sSQL & _
       "SELECT " & VarCodEntradaNovo & ", " & _
       "CONVERT(DATETIME, '" & sData & "', 103), " & _
       "CONVERT(DATETIME, '" & sHora & "', 103), " & _
       "e.NumeroNota, e.ValorNota, e.DataEmissao, e.DataSaida, e.HoraSaida, " & _
       "e.CodigoCorrentista, e.TranspCodigo, e.TranspValorServico, e.ChavedeAcesso, " & _
       "1, CASE e.TranspModFrete " & _
       "  WHEN 0 THEN '0 - EMITENTE' WHEN 1 THEN '1 - DESTINATARIO' " & _
       "  WHEN 2 THEN '2 - TERCEIROS' WHEN 3 THEN '3 - PROP.EMITENTE' " & _
       "  WHEN 4 THEN '4 - PROP.DESTINAT.' ELSE '9 - SEM TRANSPORTE' END, " & _
       "e.CodigoNota, e.SerieNF, e.ModeloNF, e.NaturezaOperacao, e.RazaoSocial, e.CNPJ_CPF, " & _
       "e.ValorProdutos, e.ValorDesconto, e.ValorIPI, e.ValorICMS, e.ValorICMSST " & _
       "FROM EntradaEstoque e WHERE e.CodigoNota = " & lCompra & ";"

msgResultado = SQLExecuta(sSQL)
If Not Vazio(msgResultado) Then
    MsgBox "Erro ao criar entrada no estoque: " & msgResultado, vbCritical, vgAtencao
    Exit Sub
End If

txtCodEntrada.Text = VarCodEntradaNovo

msgResultado = SQLExecuta("UPDATE EntradaEstoque SET Cod_entrada = " & VarCodEntradaNovo & " WHERE CodigoNota = " & lCompra)
If Not Vazio(msgResultado) Then
    MsgBox "Erro ao vincular entrada: " & msgResultado, vbCritical, vgAtencao
End If
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
cboConNotaMes.AddItem "Marao"
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
cboFiltroNota.AddItem "PERaODO"
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
    lblConNotaNumNota.Caption = "Nam. Nota"
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
ElseIf cboFiltroNota.Text = "PERaODO" Then
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
cboOrdem.AddItem "EMISSaO"
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
       If Existe(xCaminhoXML) = True Then
           If ImportarXML(xCaminhoXML) Then CriarEntradaEstoque Compra
       End If
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
Dim sURL As String
Dim sChrome As String
Dim sFirefox As String

sURL = "https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ+gAVw2g="

sChrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
If Not Existe(sChrome) Then sChrome = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

sFirefox = "C:\Program Files\Mozilla Firefox\firefox.exe"
If Not Existe(sFirefox) Then sFirefox = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"

If Existe(sChrome) Then
    ShellExecute hwnd, "open", sChrome, sURL, vbNullString, conSwNo
ElseIf Existe(sFirefox) Then
    ShellExecute hwnd, "open", sFirefox, sURL, vbNullString, conSwNo
Else
    ShellExecute hwnd, "open", sURL, vbNullString, vbNullString, conSwNo
End If
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
      REL_Prod_Entrada_NotaXML.dfTipo.Caption = "Tipo: Mas = " & cboConNotaMes.Text & "/" & cboConNotaAno.Text
   ElseIf cboFiltroNota.Text = "PERaODO" Then
      REL_Prod_Entrada_NotaXML.dfTipo.Caption = "Tipo: Datas = " & mskConNotaInicial.Text & " a " & mskConNotaFinal.Text
   ElseIf cboFiltroNota.Text = "FORNECEDOR" Then
      REL_Prod_Entrada_NotaXML.dfTipo.Caption = "Tipo: Fornecedor = " & cboConNotaCliente.Text & ""
   ElseIf cboFiltroNota.Text = "NOTA FISCAL" Then
      REL_Prod_Entrada_NotaXML.dfTipo.Caption = "Tipo: Nota Fiscal Na " & txtConNotaNumNota.Text & ""
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
'   ElseIf cboConsulta.Text = "REFERaNCIA" Then
'      REL_Prod_Entrada_PorProduto.dfTipo.Caption = "Tipo: Referancia " & cboConsDescricao.Text & ""
'   ElseIf cboConsulta.Text = "CaD. BARRA" Then
'      REL_Prod_Entrada_PorProduto.dfTipo.Caption = "Tipo: Cad. de Barra " & cboConsDescricao.Text & ""
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
    MsgBox "Nao a possivel excluir uma nota fiscal ao qual produtos ja foram adicionados aos estoque!", vbInformation, "Aviso do Sistema"
End If
End Sub

Private Sub cmdExcluirItem_Click()
If GridNotasItens.TextMatrix(GridNotasItens.Row, 19) = "SIM" Then
    MsgBox "Produto ja adicionado no estoque fisico. Nao podera ser excluado da lista!", vbInformation, "Aviso do Sistema"
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
   ShowMsg "Insira uma quantidade valida!", vbExclamation
   txtQuantFinal.SetFocus
   Exit Sub
End If

If txtCusto.Text = "" Or txtCusto.Text = "0" Then
   ShowMsg "Insira um custo de produto valido!", vbExclamation
   txtCusto.SetFocus
   Exit Sub
End If

Dim var_COD_ITENS As Long

'AUTONUMERAaaO
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
    .TextMatrix(0, 4) = "CaD."
    .TextMatrix(0, 5) = "DESCRIaaO"
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
    
   'colocar os cabeaalho em negrito
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
   
   'AUTONUMERAaaO
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

'AUTONUMERAaaO
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
   AutoNumeracao & ", " & txtCodProdExist.Text & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), " & txtCodEntrada.Text & ", 'ENTRADA', " & Replace(CDbl(txtQuantFinal.Text), ",", ".") & ", 'ADIaaO', '" & Format(Now, ocHRMN) & "', " & vCodFunc & ", " & Replace(CDbl(txtQuantExist.Text), ",", ".") & ");"
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
   'A atualizaaao deve ser feita utilizando o comando UPDATE do sql
   'e nao mais usando o matodo .Update do Recordset
   
   'Nao se deve comparar se o campo esta vazio ou nao, pois dessa forma nao
   'havera atualizaaao quando for necessario apagar alguma informaaao
   
   Dim sSQL As String
   
   'Comando de atualizaaao
   sSQL = "UPDATE setor SET setor = '" & cboSetor.Text & "' WHERE (cod_setor = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualizaaao
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
If ShowMsg("Tem certeza que deseja alterar a descriaao do produto: " & txtDescricaoExist.Text & " pela descriaao " & txtDescricaoSel.Text & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
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

'mostrar o ultimo preao de compra
sSQL = "SELECT TOP 1 venda FROM produtos_entrada_itens WHERE (codigo_produto = " & txtCodProduto & ") ORDER BY codigo DESC;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then vrVenda = r("venda")
If r.State <> 0 Then r.Close
Set r = Nothing

txtValor.Text = Format(vrVenda, ocMONEY)
End Sub

Private Sub cmdCancelarEntrada_Click()
lblProdutoCadastrado.Caption = "Esse produto nao foi encontrado no seu banco de dados!"
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

varData = Empty                    'Inicializa a variavel

Set fCal = New Calendario      'Cria o form de calendario
fCal.Show vbModal

varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrai a variavel

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskConNotaInicial = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdConNotaCal2_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty                    'Inicializa a variavel

Set fCal = New Calendario      'Cria o form de calendario
fCal.Show vbModal

varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrai a variavel

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskConNotaFinal = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub


Private Sub cmdExibirConNotas_Click()
Dim totalRegistros As Long

On Error GoTo ErrLoad

'INDICE
Dim INDICE As String
If cboOrdem.Text = "EMISSaO" Then
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
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 " & _
                    "ORDER BY " & INDICE
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
                        "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                        "FROM EntradaEstoque AS EntradaEstoque_1 " & _
                        "ORDER BY " & INDICE
ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
    If txtConNotaNumNota.Text = "" Then Exit Sub
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 WHERE NumeroNota = " & txtConNotaNumNota & " " & _
                    "ORDER BY " & INDICE
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
                        "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                        "FROM EntradaEstoque AS EntradaEstoque_1 WHERE NumeroNota = " & txtConNotaNumNota & " " & _
                        "ORDER BY " & INDICE
ElseIf cboFiltroNota.Text = "CHAVE DE ACESSO" Then
    If txtConNotaNumNota.Text = "" Then Exit Sub
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, ChavedeAcesso," & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 WHERE ChavedeAcesso = '" & txtConNotaNumNota & "' " & _
                    "ORDER BY " & INDICE
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, ChavedeAcesso, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
                        "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                        "FROM EntradaEstoque AS EntradaEstoque_1 WHERE ChavedeAcesso = '" & txtConNotaNumNota & "' " & _
                        "ORDER BY " & INDICE
                        'Debug.Print sSQL
ElseIf cboFiltroNota.Text = "FORNECEDOR" Then
    If txtConNotaCodCliente.Text = "" Then Exit Sub
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 WHERE CodigoCorrentista = " & txtConNotaCodCliente.Text & " " & _
                    "ORDER BY " & INDICE
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
                        "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                        "FROM EntradaEstoque AS EntradaEstoque_1 WHERE CodigoCorrentista = " & txtConNotaCodCliente.Text & " " & _
                        "ORDER BY " & INDICE
ElseIf cboFiltroNota.Text = "PERaODO" Then
    If mskConNotaInicial.Text = "" Then Exit Sub
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 WHERE (DataEmissao >= CONVERT(DATETIME, '" & Format(mskConNotaInicial.Text, ocDATA) & "', 103)) AND (DataEmissao <= CONVERT(DATETIME, '" & Format(mskConNotaFinal.Text, ocDATA) & "', 103)) " & _
                    "ORDER BY " & INDICE
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
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
    ElseIf cboConNotaMes.Text = "marao" Then
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
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                    "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
                    "(SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) AS vSomaPendencia  " & _
                    "FROM EntradaEstoque AS EntradaEstoque_1 WHERE " & vWhere & " AND (YEAR(DataEmissao) = " & cboConNotaAno & ") order by CodigoNota desc"
    sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, RazaoSocial, ValorNota, " & _
                        "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'RECEBIDA' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status, " & _
                        "(CASE WHEN (SELECT COUNT(Item) FROM EntradaEstoqueItens WHERE (CodigoNota = EntradaEstoque_1.CodigoNota) AND (Adicionada = 0)) < 1 THEN ' ' ELSE 'PENDaNCIA' END) AS SITUACAO,  " & _
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

Private Sub cmdVinculos_Click()
   Dim lCodNota As Long
   lCodNota = Val(txtCodNota.Text)
   If lCodNota = 0 Then
      MsgBox "Nenhuma nota selecionada.", vbExclamation
      Exit Sub
   End If
   Dim frmVinc As New frmVinculoProdutoXML
   frmVinc.NumeroEntrada = lCodNota
   frmVinc.MostrarTodos  = True
   frmVinc.Show vbModal, Me
   Unload frmVinc
   Set frmVinc = Nothing
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
 .DialogTitle = "Selecionar o arquivo XML para importaaao"  'tiulo do dilogo...
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
        If ImportarXML(arqXML) Then
            CriarEntradaEstoque Compra
            cmdExcluir.Enabled = True
        End If
    End If
 End If
End With
Exit Sub

deuErro:
 Set dlgArquivo = Nothing
 MsgBox Err.Description, vbCritical, vgAtencao
 Err.Clear
End Sub

Public Function ImportarXML(arqXML As String) As Boolean
Dim CNPJ As String, CPF As String, tipoEmp As Integer
Dim cstpis As String, cstcofins As String, csticms As String, cstipi As String, cEAN As String
Dim i As Integer
Dim qtdProd As Integer, qtdDup As Integer, xCfop As Integer, qtfracao As Integer
Dim CodProduto As Long, Codigo_fornecedor As Long

'--- Variaveis ICMS por item ---
Dim subgrupoICMS As String       'Nome do subgrupo detectado: ICMS00, ICMS10, ICMS20, etc.
Dim ICMSBase As Double           'vBC:      Base de calculo do ICMS
Dim ICMSAliquota As Double       'pICMS:    Aliquota do ICMS (%)
Dim ICMSValor As Double          'vICMS:    Valor do ICMS calculado
Dim ICMSmodBC As Integer         'modBC:    Modalidade da base de calculo (0=Margem Valor Agregado, 1=Pauta, 2=Preco Tabelado Max, 3=Valor Operacao)
Dim ICMSpRedBC As Double         'pRedBC:   Percentual de reducao da base de calculo
Dim ICMSvICMSOp As Double        'vICMSOp:  Valor do ICMS da Operacao (ICMS51 - diferimento)
Dim ICMSpDif As Double           'pDif:     Percentual do diferimento (ICMS51)
Dim ICMSvICMSDeson As Double     'vICMSDeson: Valor do ICMS desonerado (ICMS40/41)
Dim ICMSmotDesICMS As String     'motDesICMS: Motivo da desoneracao do ICMS (ICMS40/41)
'--- Variaveis ICMS-ST (Substituicao Tributaria) ---
Dim ICMSmodBCST As Integer       'modBCST:  Modalidade da base de calculo do ICMS ST
Dim ICMSpRedBCST As Double       'pRedBCST: Percentual de reducao da base de calculo ST
Dim ICMSpMVAST As Double         'pMVAST:   Percentual de Margem de Valor Adicionado ST
Dim ICMSvBCST As Double          'vBCST:    Valor da base de calculo do ICMS ST
Dim ICMSpICMSST As Double        'pICMSST:  Aliquota do ICMS ST (%)
Dim ICMSvICMSST As Double        'vICMSST:  Valor do ICMS ST
Dim ICMSvBCFCPST As Double       'vBCFCPST: Base de calculo do FCP retido por ST
Dim ICMSpFCPST As Double         'pFCPST:   Percentual do FCP retido por ST
Dim ICMSvFCPST As Double         'vFCPST:   Valor do FCP retido por ST
'--- Variaveis ICMS-ST Retido (ICMS60) ---
Dim ICMSvBCSTRet As Double       'vBCSTRet:      Base de calculo do ICMS ST retido anteriormente
Dim ICMSpST As Double            'pST:           Aliquota suportada pelo consumidor final
Dim ICMSvICMSSubstituto As Double 'vICMSSubstituto: Valor do ICMS do substituto
Dim ICMSvICMSSTRet As Double     'vICMSSTRet:    Valor do ICMS ST retido anteriormente
Dim ICMSvBCFCPSTRet As Double    'vBCFCPSTRet:   Base de calculo do FCP ST retido
Dim ICMSpFCPSTRet As Double      'pFCPSTRet:     Percentual do FCP ST retido
Dim ICMSvFCPSTRet As Double      'vFCPSTRet:     Valor do FCP ST retido

'--- Variaveis IPI por item ---
Dim IPIcEnq As String            'cEnq:  Codigo de enquadramento legal do IPI
Dim IPIvBC As Double             'vBC:   Base de calculo do IPI
Dim IPIpIPI As Double            'pIPI:  Aliquota do IPI (%)
Dim IPIvIPI As Double            'vIPI:  Valor do IPI

'--- Variaveis PIS por item ---
Dim pisCST As String             'CST:       Codigo de Situacao Tributaria do PIS
Dim PISvBC As Double             'vBC:       Base de calculo do PIS (CST 01/02)
Dim PISpPIS As Double            'pPIS:      Aliquota do PIS em percentual (CST 01/02)
Dim PISvPIS As Double            'vPIS:      Valor do PIS (CST 01/02/03/99)
Dim PISqBCProd As Double         'qBCProd:   Quantidade vendida (CST 03 - PISQtde)
Dim PISvAliqProd As Double       'vAliqProd: Aliquota do PIS em reais por unidade (CST 03)

'--- Variaveis COFINS por item ---
Dim cofinsCST As String          'CST:       Codigo de Situacao Tributaria do COFINS
Dim COFINSvBC As Double          'vBC:       Base de calculo do COFINS (CST 01/02)
Dim COFINSpCOFINS As Double      'pCOFINS:   Aliquota do COFINS em percentual (CST 01/02)
Dim COFINSvCOFINS As Double      'vCOFINS:   Valor do COFINS (CST 01/02/03/99)
Dim COFINSqBCProd As Double      'qBCProd:   Quantidade vendida (CST 03 - COFINSQtde)
Dim COFINSvAliqProd As Double    'vAliqProd: Aliquota do COFINS em reais por unidade (CST 03)'
Dim Fornecedores As New ADODB.Recordset, Compras_XML As New ADODB.Recordset, TBVerificarNFe As New ADODB.Recordset, Compras_Itens_XML As New ADODB.Recordset, Produtos As New ADODB.Recordset
Dim Produtos_fornecedores As New ADODB.Recordset
Dim TbTransportadora As New ADODB.Recordset
Dim Codigo_transportadora As Long
Dim varCNPJtransp As String, varNomeTransp As String
Dim varIEtransp As String, varEndtransp As String, varCidtransp As String, varUFtransp As String
Dim varTipotransp As String, varCodTransp As Long
Dim sDupNum As String, sDupVenc As String, vDupValor As Double, sFavorecido As String
Dim TbVinculo As New ADODB.Recordset   'Vinculo XML: mapeamento embalagem fornecedor -> produto interno
Dim xCProd As String                   'cProd do item atual da XML
Dim xVUnCom As Double                  'vUnCom do item atual da XML (preco unitario comercial do fornecedor)

Dim XMLDOC As MSXML2.DOMDocument30
Set XMLDOC = New MSXML2.DOMDocument30
XMLDOC.Load (arqXML)

Dim str_sql As String
Dim str_sql_item As String

'--- Verificacao antecipada: NF-e ja importada? (ChavedeAcesso e unica por NF-e) ---
sSQL = "SELECT COUNT(1) r FROM EntradaEstoque WHERE ChavedeAcesso = '" & _
       XMLDOC.getElementsByTagName("nfeProc/protNFe/infProt/chNFe").item(0).Text & "'"
If SQLExecutaRetorno(sSQL, "r", 0) > 0 Then
    MsgBox "Essa Nota Fiscal ja foi adicionada anteriormente!", vbInformation, "Aviso do Sistema"
    ImportarXML = False: Exit Function
End If

On Error Resume Next

'-------- IMPORTANDO OS DADOS DO FORNECEDOR ---------

'Trata da Mascara Cnpj\cpf
'CNPJ = Retira(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/CNPJ").Item(0).Text, "./-", UM_A_UM) 'desmarquei
CNPJ = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/CNPJ").item(0).Text

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



'Verifica se o Fornecedor esta cadastrado ou nao
If Fornecedores.RecordCount = 0 Then
    'Fornecedor nao encontrado: cadastra automaticamente com os dados do emitente da NF-e
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
    Dim varTelefoneForn As String
    Dim varComplementoForn As String

    varCodFornecedor = SQLExecutaRetorno("SELECT ISNULL(MAX(Codigo),0) r FROM fornecedor", "r", 0) + 1
    'Replace para escapar aspas simples no nome â€” evita erro SQL em razoes sociais com apostrofo
    varRAZAOForn = Replace(Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xNome").item(0).Text), 35), "'", "''")
    varFantasiaFor = Replace(Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xFant").item(0).Text), 35), "'", "''")
    If tipoEmp = 0 Then varCNPJforn = CPF
    If tipoEmp = 1 Then varCNPJforn = CNPJ
    varIEforn = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/IE").item(0).Text
    varENDERECOforn = Replace(Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xLgr").item(0).Text), 40), "'", "''")
    varNUMforn = Val(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").item(0).Text)
    varBairroforn = Replace(Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xBairro").item(0).Text), 20), "'", "''")
    varCidadeforn = Replace(Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xMun").item(0).Text), 20), "'", "''")
    varEstadoforn = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").item(0).Text)
    varCEPform = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/CEP").item(0).Text
    varCODIBGEform = Val(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/cMun").item(0).Text)

    'Telefone: tag <fone> em enderEmit (opcional)
    If XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/fone").Length > 0 Then
       varTelefoneForn = XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/fone").item(0).Text
    End If

    'Complemento: tag <xCpl> em enderEmit (opcional)
    If XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/xCpl").Length > 0 Then
       varComplementoForn = Replace(Left(UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/xCpl").item(0).Text), 30), "'", "''")
    End If

    'Comando de inclusao
    sSQL = "INSERT INTO fornecedor (" & _
           "CODIGO, RAZAO, FANTASIA, CPF, IE, ENDERECO, Numero, Bairro, Complemento, Cidade, Estado, CEP, CodigoIBGE, Telefone" & _
           ") VALUES (" & _
           varCodFornecedor & ", " & _
           "'" & varRAZAOForn & "', " & _
           "'" & varFantasiaFor & "', " & _
           "'" & varCNPJforn & "', " & _
           "'" & varIEforn & "', " & _
           "'" & varENDERECOforn & "', " & _
           varNUMforn & ", " & _
           "'" & varBairroforn & "', " & _
           "'" & varComplementoForn & "', " & _
           "'" & varCidadeforn & "', " & _
           "'" & varEstadoforn & "', " & _
           "'" & varCEPform & "', " & _
           varCODIBGEform & ", " & _
           "'" & varTelefoneForn & "')"

    dbData.Execute (sSQL)

    'Define Codigo_fornecedor para uso no restante do sub (CodigoCorrentista, a_pagar, etc.)
    Codigo_fornecedor = varCodFornecedor

Else
    'Fornecedor ja cadastrado: atualiza endereco com os dados mais recentes da NF-e
    Codigo_fornecedor = Fornecedores!Codigo
    Fornecedores!Endereco = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xLgr").item(0).Text), 40)
    Fornecedores!Numero = Val(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").item(0).Text)
    Fornecedores!bairro = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xBairro").item(0).Text), 20)
    Fornecedores!Cidade = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xMun").item(0).Text), 20)
    Fornecedores!Estado = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").item(0).Text)
    Fornecedores!CEP = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/CEP").item(0).Text
    If XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/xCpl").Length > 0 Then
       Fornecedores!Complemento = Left(UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/xCpl").item(0).Text), 30)
    End If
    If XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/fone").Length > 0 Then
       Fornecedores!telefone = XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/fone").item(0).Text
    End If
    Fornecedores.Update
End If
'-------- FIM DA IMPORTAaaO DO FORNECEDOR --------

'-------- IMPORTANDO DADOS DA TRANSPORTADORA ---------
Codigo_transportadora = 0
If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta").Length > 0 Then

   'Obtem CNPJ ou CPF da transportadora (mutuamente exclusivos na NF-e)
   varCNPJtransp = ""
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/CNPJ").Length > 0 Then
      varCNPJtransp = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/CNPJ").item(0).Text
      varTipotransp = "PJ"  'Pessoa Juridica
   ElseIf XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/CPF").Length > 0 Then
      varCNPJtransp = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/CPF").item(0).Text
      varTipotransp = "PF"  'Pessoa Fisica
   End If

   If Len(varCNPJtransp) > 0 Then
      RsOpen TbTransportadora, "SELECT * FROM Transportadora WHERE CNPJ = '" & varCNPJtransp & "'"

      If TbTransportadora.RecordCount = 0 Then
         'Transportadora nao encontrada: cadastra automaticamente com dados da NF-e
         varCodTransp = SQLExecutaRetorno("SELECT ISNULL(MAX(CODIGO),0) r FROM Transportadora", "r", 0) + 1
         varNomeTransp = Replace(Left(UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/xNome").item(0).Text), 35), "'", "''")
         varIEtransp = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/IE").item(0).Text
         varEndtransp = Replace(Left(UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/xEnder").item(0).Text), 40), "'", "''")
         varCidtransp = Replace(Left(UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/xMun").item(0).Text), 20), "'", "''")
         varUFtransp = UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/UF").item(0).Text)

         sSQL = "INSERT INTO Transportadora " & _
                "(CODIGO, RAZAO, CNPJ, IE, Endereco, Cidade, Estado, TIPO) VALUES (" & _
                varCodTransp & ", " & _
                "'" & varNomeTransp & "', " & _
                "'" & varCNPJtransp & "', " & _
                "'" & varIEtransp & "', " & _
                "'" & varEndtransp & "', " & _
                "'" & varCidtransp & "', " & _
                "'" & varUFtransp & "', " & _
                "'" & varTipotransp & "')"

         dbData.Execute (sSQL)
         Codigo_transportadora = varCodTransp

      Else
         'Transportadora ja cadastrada: atualiza endereco com dados mais recentes da NF-e
         Codigo_transportadora = TbTransportadora!Codigo
         TbTransportadora!Endereco = Replace(Left(UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/xEnder").item(0).Text), 40), "'", "''")
         TbTransportadora!Cidade = Left(UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/xMun").item(0).Text), 20)
         TbTransportadora!Estado = UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/UF").item(0).Text)
         TbTransportadora!IE = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/IE").item(0).Text
         TbTransportadora.Update
      End If

      TbTransportadora.Close
   End If
End If
'-------- FIM DA IMPORTACAO DA TRANSPORTADORA --------

   '--------CADASTRO DA ENTRADA NO ESTOQUE ---------
   RsOpen Compras_XML, "SELECT TOP 1 * FROM EntradaEstoque"
   
   Compra = SQLExecutaRetorno("SELECT ISNULL(MAX(CodigoNota),0) r FROM EntradaEstoque", "r", 0) + 1
   
   Compras_XML.AddNew
   Compras_XML!CodigoNota = Compra
   Compras_XML!ModeloNF = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/mod").item(0).Text
   Compras_XML!NumeroNota = Format(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/nNF").item(0).Text, "000000000")
   Compras_XML!SerieNF = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/serie").item(0).Text
   Compras_XML!ChavedeAcesso = XMLDOC.getElementsByTagName("nfeProc/protNFe/infProt/chNFe").item(0).Text
   Compras_XML!NaturezaOperacao = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/natOp").item(0).Text)
   
   If XMLDOC.getElementsByTagName("nfeProc/protNFe").item(0).Attributes.getNamedItem("versao").Text = "2.00" Then
       Compras_XML!DataEmissao = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/dEmi").item(0).Text
       Compras_XML!DataSaida = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/dEmi").item(0).Text
   Else
       Compras_XML!DataEmissao = Left(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/dhEmi").item(0).Text, 10)
       Compras_XML!DataSaida = Left(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/dhEmi").item(0).Text, 10)
   End If
  
   'Compras_XML!TipoEntradaEst = "Nota Fiscal"
   Compras_XML!HoraSaida = Time
   
   'Codigo_fornecedor ja definido em ambas as branches do bloco acima (novo ou existente)
   Compras_XML!CodigoCorrentista = Codigo_fornecedor
       
   Compras_XML!RazaoSocial = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xNome").item(0).Text)
   Compras_XML!InscricaoEstadual = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/IE").item(0).Text)
   
   'Compras_XML!CNPJ_CPF = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/CNPJ").Item(0).Text)
   If tipoEmp = 0 Then Compras_XML!CNPJ_CPF = CPF
   If tipoEmp = 1 Then Compras_XML!CNPJ_CPF = CNPJ
   
   Compras_XML!logradouro = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xLgr").item(0).Text)
   Compras_XML!Numero = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").item(0).Text)
   'xCpl: Complemento do endereco do emitente - tag <xCpl> em enderEmit (opcional)
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/xCpl").Length > 0 Then
      Compras_XML!Complemento = UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/xCpl").item(0).Text)
   End If
   Compras_XML!bairro = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xBairro").item(0).Text)
   Compras_XML!CodigoIBGE = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/cMun").item(0).Text)
   Compras_XML!Municipio = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xMun").item(0).Text)
   Compras_XML!UF = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").item(0).Text)
   Compras_XML!CEP = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/CEP").item(0).Text)
   Compras_XML!telefone = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/fone").item(0).Text)

   ' indPag: em NF-e 4.0+ esta em pag/detPag/indPag; em NF-e 3.x estava em ide/indPag (removido)
   ' sIndPag vazio = tag ausente na XML (NF-e antigas sem o campo)
   Dim sIndPag As String
   sIndPag = ""
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/pag/detPag/indPag").Length > 0 Then
      sIndPag = XMLDOC.selectNodes("nfeProc/NFe/infNFe/pag/detPag/indPag").item(0).Text
   ElseIf XMLDOC.selectNodes("nfeProc/NFe/infNFe/ide/indPag").Length > 0 Then
      sIndPag = XMLDOC.selectNodes("nfeProc/NFe/infNFe/ide/indPag").item(0).Text
   End If
   Select Case sIndPag
      Case "0"
         Compras_XML!IndicadorFormaPagamento = "0 - Pagamento a vista"
      Case "1"
         Compras_XML!IndicadorFormaPagamento = "1 - Pagamento a prazo"
      Case "2"
         Compras_XML!IndicadorFormaPagamento = "2 - Outros"
      Case Else
         Compras_XML!IndicadorFormaPagamento = "2 - Outros"
   End Select
   ' vPag: Valor do pagamento declarado pelo emitente - tag <vPag> dentro de <pag/detPag>
   Compras_XML!PagvPag = 0
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/pag/detPag/vPag").Length > 0 Then
      Compras_XML!PagvPag = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/pag/detPag/vPag").item(0).Text, ".", ",", UM_A_UM))
   End If

   Select Case Int(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/tpImp").item(0).Text)
      Case 1
          Compras_XML!FormatoImpressaoDANFE = "1 - Retrato"
      Case 2
          Compras_XML!FormatoImpressaoDANFE = "2 - Paisagem"
   End Select
   
   Select Case Int(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/tpEmis").item(0).Text)
      Case 1
          Compras_XML!FormatoEmissaoNFe = "1 - Normal"
      Case 2
          Compras_XML!FormatoEmissaoNFe = "2 - Contingancia FS"
      Case 3
          Compras_XML!FormatoEmissaoNFe = "3 - Contingancia SCAN"
      Case 4
          Compras_XML!FormatoEmissaoNFe = "4 - Contingancia DPEC"
      Case 5
          Compras_XML!FormatoEmissaoNFe = "5 - Contingancia FS-DA"
   End Select
   
   'idDest: Identificador de Local de destino - 1=Op Interna, 2=Op Interestadual, 3=Op c/ Exterior - tag <idDest> em ide
   Compras_XML!IdentificadorDestino = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/idDest").item(0).Text
   Compras_XML!CodigoNatureza = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/idDest").item(0).Text
   'indPres: Indicador de Presenca do Comprador - 0=Nao se aplica, 1=Op presencial, 2=Nao presencial Internet etc - tag <indPres> em ide
   Compras_XML!IndicadorPresencaComprador = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/indPres").item(0).Text
   'indIEDest: Indicador da IE do Destinatario - 1=Contribuinte ICMS, 2=Contribuinte isento, 9=Nao contribuinte - tag <indIEDest> em dest
   Compras_XML!IndicadorIEDestinatario = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/dest/indIEDest").item(0).Text
   'finNFe: Finalidade de emissao - 1=Normal, 2=Complementar, 3=Ajuste, 4=Devolucao/Retorno - tag <finNFe> em ide
   Compras_XML!FinalidadeEmissaoNFe = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/finNFe").item(0).Text
   
   Compras_XML!modFrete = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/transp/modFrete").item(0).Text
   
   If Int(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/infAdic/infCpl").Length) > 0 Then
      Compras_XML!InformacoesAdicionais = Left(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/infAdic/infCpl").item(0).Text, 100)
   End If
   
   Compras_XML!BaseICMS = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vBC").item(0).Text, ".", ",", UM_A_UM))         'vBC:        Base de calculo do ICMS
   Compras_XML!ValorICMS = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vICMS").item(0).Text, ".", ",", UM_A_UM))       'vICMS:      Valor total do ICMS
   Compras_XML!ValorIsento = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vICMSDeson").item(0).Text, ".", ",", UM_A_UM)) 'vICMSDeson: ICMS desonerado
   Compras_XML!ValorFCP = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vFCP").item(0).Text, ".", ",", UM_A_UM))         'vFCP:       FCP sobre ICMS normal
   Compras_XML!BaseICMSST = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vBCST").item(0).Text, ".", ",", UM_A_UM))      'vBCST:      Base de calculo do ICMS ST
   Compras_XML!ValorICMSST = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vST").item(0).Text, ".", ",", UM_A_UM))       'vST:        Valor total do ICMS ST
   Compras_XML!ValorFCPST = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vFCPST").item(0).Text, ".", ",", UM_A_UM))     'vFCPST:     FCP retido por ST
   Compras_XML!ValorFCPSTRet = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vFCPSTRet").item(0).Text, ".", ",", UM_A_UM)) 'vFCPSTRet: FCP retido anteriormente por ST
   Compras_XML!ValorProdutos = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vProd").item(0).Text, ".", ",", UM_A_UM))   'vProd:      Valor total dos produtos
   Compras_XML!ValorFrete = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vFrete").item(0).Text, ".", ",", UM_A_UM))     'vFrete:     Valor total do frete
   Compras_XML!ValorSeguro = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vSeg").item(0).Text, ".", ",", UM_A_UM))      'vSeg:       Valor total do seguro
   Compras_XML!ValorDesconto = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vDesc").item(0).Text, ".", ",", UM_A_UM))   'vDesc:      Valor total do desconto
   Compras_XML!ValorImportacao = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vII").item(0).Text, ".", ",", UM_A_UM))   'vII:        Valor total do Imposto de Importacao
   Compras_XML!ValorIPI = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vIPI").item(0).Text, ".", ",", UM_A_UM))         'vIPI:       Valor total do IPI
   Compras_XML!ValorIPIDevol = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vIPIDevol").item(0).Text, ".", ",", UM_A_UM)) 'vIPIDevol: Valor total do IPI devolvido
   Compras_XML!ValorPIS = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vPIS").item(0).Text, ".", ",", UM_A_UM))         'vPIS:       Valor total do PIS
   Compras_XML!ValorCOFINS = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vCOFINS").item(0).Text, ".", ",", UM_A_UM))   'vCOFINS:    Valor total do COFINS
   Compras_XML!ValorOutrasDespesas = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vOutro").item(0).Text, ".", ",", UM_A_UM)) 'vOutro:  Outras despesas acessorias
   Compras_XML!ValorNota = CDbl(Substitui(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/total/ICMSTot/vNF").item(0).Text, ".", ",", UM_A_UM))         'vNF:        Valor total da NF-e

   '--- ICMS DIFAL - Diferencial de Aliquota (EC 87/2015 / NT 2015/003 - operacoes B2C interestaduais - tags opcionais) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/vFCPUFDest").Length > 0 Then
      Compras_XML!vFCPUFDest = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/vFCPUFDest").item(0).Text, ".", ",", UM_A_UM))    'vFCPUFDest: FCP p/ UF de destino
      Compras_XML!vICMSUFDest = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/vICMSUFDest").item(0).Text, ".", ",", UM_A_UM))  'vICMSUFDest: ICMS p/ UF de destino
      Compras_XML!vICMSUFRemet = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/vICMSUFRemet").item(0).Text, ".", ",", UM_A_UM)) 'vICMSUFRemet: ICMS p/ UF de origem
   End If

   '--- ICMS Monofasico (NF-e 4.00+ NT 2021/004 - combustiveis e monofasicos - tags opcionais) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/qBCMono").Length > 0 Then
      Compras_XML!qBCMono = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/qBCMono").item(0).Text, ".", ",", UM_A_UM))          'qBCMono:     Qtde BC ICMS Monofasico
      Compras_XML!vICMSMono = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/vICMSMono").item(0).Text, ".", ",", UM_A_UM))      'vICMSMono:   Valor ICMS Monofasico
   End If
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/qBCMonoReten").Length > 0 Then
      Compras_XML!qBCMonoReten = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/qBCMonoReten").item(0).Text, ".", ",", UM_A_UM))      'qBCMonoReten:  Qtde BC Mono retido por ST
      Compras_XML!vICMSMonoReten = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/vICMSMonoReten").item(0).Text, ".", ",", UM_A_UM))  'vICMSMonoReten:Valor Mono retido por ST
   End If
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/qBCMonoRet").Length > 0 Then
      Compras_XML!qBCMonoRet = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/qBCMonoRet").item(0).Text, ".", ",", UM_A_UM))      'qBCMonoRet:  Qtde BC Mono retido anteriormente
      Compras_XML!vICMSMonoRet = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/vICMSMonoRet").item(0).Text, ".", ",", UM_A_UM))  'vICMSMonoRet:Valor Mono retido anteriormente
   End If
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/vTotTrib").Length > 0 Then
      Compras_XML!TotTributos = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ICMSTot/vTotTrib").item(0).Text, ".", ",", UM_A_UM))      'vTotTrib:    Total aprox. tributos da NF-e (IBPT)
   End If

   Compras_XML!NumeroProtocolo = XMLDOC.getElementsByTagName("nfeProc/protNFe/infProt/nProt").item(0).Text
   'dhRecbto: Data e hora do recebimento/processamento da NF-e na SEFAZ - tag <dhRecbto> em infProt
   Compras_XML!DataHoraRecebimento = XMLDOC.getElementsByTagName("nfeProc/protNFe/infProt/dhRecbto").item(0).Text
   Compras_XML!DataHoraProcotolo = XMLDOC.getElementsByTagName("nfeProc/protNFe/infProt/dhRecbto").item(0).Text
   Compras_XML!enviada = 1
   'Compras_XML!ImportadoXML = True

   '--- Campos de identificacao adicionais ---

   'indFinal: Indicador de operacao com Consumidor Final - 0=Normal, 1=Consumidor Final - tag <indFinal> em ide
   Compras_XML!ConsumidorFinal = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/ide/indFinal").item(0).Text

   'IM: Inscricao Municipal do emitente - tag <IM> em emit (opcional)
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/IM").Length > 0 Then
      Compras_XML!InscricaoMunicipal = XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/IM").item(0).Text
   End If

   'SUFRAMA: Codigo SUFRAMA do emitente - tag <SUFRAMA> em emit (apenas para Zona Franca de Manaus)
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/SUFRAMA").Length > 0 Then
      Compras_XML!Suframa = XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/SUFRAMA").item(0).Text
   End If

   'cPais/xPais: Codigo e nome do pais do endereco do emitente - tags <cPais> e <xPais> em enderEmit (opcional)
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/cPais").Length > 0 Then
      Compras_XML!CodigoPais = XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/cPais").item(0).Text
      Compras_XML!Pais = XMLDOC.selectNodes("nfeProc/NFe/infNFe/emit/enderEmit/xPais").item(0).Text
   End If

   'email: Email do destinatario - tag <email> em dest (opcional)
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/dest/email").Length > 0 Then
      Compras_XML!email = XMLDOC.selectNodes("nfeProc/NFe/infNFe/dest/email").item(0).Text
   End If

   'infAdFisco: Informacoes adicionais de interesse do Fisco - tag <infAdFisco> em infAdic (opcional)
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/infAdic/infAdFisco").Length > 0 Then
      Compras_XML!InformacoesComplementares = Left(XMLDOC.selectNodes("nfeProc/NFe/infNFe/infAdic/infAdFisco").item(0).Text, 100)
   End If

   'dhCont/xJust: Data/hora e justificativa de entrada em contingencia - tags <dhCont> e <xJust> em ide (apenas em contingencia)
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/ide/dhCont").Length > 0 Then
      Compras_XML!ContingenciaDataHora = XMLDOC.selectNodes("nfeProc/NFe/infNFe/ide/dhCont").item(0).Text
      Compras_XML!ContingenciaJustificativa = XMLDOC.selectNodes("nfeProc/NFe/infNFe/ide/xJust").item(0).Text
   End If

   'UFEmbar/xLocEmbar: UF e local de embarque - tags <UFEmbar> e <xLocEmbar> em exporta (apenas exportacao)
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/exporta").Length > 0 Then
      Compras_XML!UFEmbarque = XMLDOC.selectNodes("nfeProc/NFe/infNFe/exporta/UFEmbar").item(0).Text
      Compras_XML!LocEmbarque = XMLDOC.selectNodes("nfeProc/NFe/infNFe/exporta/xLocEmbar").item(0).Text
   End If

   '--- Totais ISSQNtot (presentes apenas em NF-e com servicos sujeitos ao ISS) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot").Length > 0 Then
      'vServ: Valor total dos servicos sob nao incidencia ou nao tributados pelo ICMS - tag <vServ>
      Compras_XML!ValorServico = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vServ").item(0).Text, ".", ",", UM_A_UM))
      'vBC: Base de Calculo do ISS - tag <vBC>
      Compras_XML!ValorBCISS = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vBC").item(0).Text, ".", ",", UM_A_UM))
      'vISS: Valor Total do ISS - tag <vISS>
      Compras_XML!ValorISS = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vISS").item(0).Text, ".", ",", UM_A_UM))
      'vPIS: Valor do PIS sobre servicos - tag <vPIS>
      Compras_XML!ValorPISISS = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vPIS").item(0).Text, ".", ",", UM_A_UM))
      'vCOFINS: Valor do COFINS sobre servicos - tag <vCOFINS>
      Compras_XML!ValorPISCOFINS = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vCOFINS").item(0).Text, ".", ",", UM_A_UM))
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vRetPIS").Length > 0 Then
         'vRetPIS: Valor do PIS Retido - tag <vRetPIS>
         Compras_XML!ValorRetidoPIS = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vRetPIS").item(0).Text, ".", ",", UM_A_UM))
         'vRetCOFINS: Valor do COFINS Retido - tag <vRetCOFINS>
         Compras_XML!ValorRetidoCOFINS = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vRetCOFINS").item(0).Text, ".", ",", UM_A_UM))
         'vRetCSLL: Valor da CSLL Retida - tag <vRetCSLL>
         Compras_XML!ValorRetidoCSLL = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vRetCSLL").item(0).Text, ".", ",", UM_A_UM))
         'vBCIRRF: Base de Calculo do IRRF - tag <vBCIRRF>
         Compras_XML!ValorBCIRRF = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vBCIRRF").item(0).Text, ".", ",", UM_A_UM))
         'vIRRF: Valor do IRRF Retido - tag <vIRRF>
         Compras_XML!ValorIRRF = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vIRRF").item(0).Text, ".", ",", UM_A_UM))
         'vBCRetPrev: Base de Calculo da Retencao da Previdencia Social - tag <vBCRetPrev>
         Compras_XML!ValorRetidoBCPrev = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vBCRetPrev").item(0).Text, ".", ",", UM_A_UM))
         'vRetPrev: Valor da Retencao da Previdencia Social - tag <vRetPrev>
         Compras_XML!ValorRetidoPrev = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISSQNtot/vRetPrev").item(0).Text, ".", ",", UM_A_UM))
      End If
   End If

   '--- Totais IBSCBSTot (presentes apenas em NF-e com Reforma Tributaria - EC 132/2023 + LC 214/2025) ---
   '--- Suporta dois schemas: NT inicial (gIBSTot/gCBSTot/<vBC>) e NT 2024+ (gIBS/gCBS/<vBCIBSCBS>) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot").Length > 0 Then
      ' Base de calculo: <vBCIBSCBS> (NT 2024+) ou <vBC> (schema anterior)
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/vBCIBSCBS").Length > 0 Then
         Compras_XML!TotIBSCBSvBC = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/vBCIBSCBS").item(0).Text, ".", ",", UM_A_UM))
      ElseIf XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/vBC").Length > 0 Then
         Compras_XML!TotIBSCBSvBC = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/vBC").item(0).Text, ".", ",", UM_A_UM))
      End If
      ' Grupo IBS: <gIBS> (NT 2024+) ou <gIBSTot> (schema anterior)
      Dim sIBSPfx As String
      sIBSPfx = ""
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/gIBS").Length > 0 Then
         sIBSPfx = "nfeProc/NFe/infNFe/total/IBSCBSTot/gIBS"
      ElseIf XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/gIBSTot").Length > 0 Then
         sIBSPfx = "nfeProc/NFe/infNFe/total/IBSCBSTot/gIBSTot"
      End If
      If sIBSPfx <> "" Then
         ' IBS UF
         Compras_XML!TotIBSUFvDif = CDbl(Substitui(XMLDOC.selectNodes(sIBSPfx & "/gIBSUF/vDif").item(0).Text, ".", ",", UM_A_UM))
         Compras_XML!TotIBSUFvDevTrib = CDbl(Substitui(XMLDOC.selectNodes(sIBSPfx & "/gIBSUF/vDevTrib").item(0).Text, ".", ",", UM_A_UM))
         ' Valor IBS UF: <vIBSUF> (NT 2024+) ou <vIBS> (schema anterior)
         If XMLDOC.selectNodes(sIBSPfx & "/gIBSUF/vIBSUF").Length > 0 Then
            Compras_XML!TotIBSUFvIBS = CDbl(Substitui(XMLDOC.selectNodes(sIBSPfx & "/gIBSUF/vIBSUF").item(0).Text, ".", ",", UM_A_UM))
         ElseIf XMLDOC.selectNodes(sIBSPfx & "/gIBSUF/vIBS").Length > 0 Then
            Compras_XML!TotIBSUFvIBS = CDbl(Substitui(XMLDOC.selectNodes(sIBSPfx & "/gIBSUF/vIBS").item(0).Text, ".", ",", UM_A_UM))
         End If
         ' IBS Municipal
         Compras_XML!TotIBSMunvDif = CDbl(Substitui(XMLDOC.selectNodes(sIBSPfx & "/gIBSMun/vDif").item(0).Text, ".", ",", UM_A_UM))
         Compras_XML!TotIBSMunvDevTrib = CDbl(Substitui(XMLDOC.selectNodes(sIBSPfx & "/gIBSMun/vDevTrib").item(0).Text, ".", ",", UM_A_UM))
         ' Valor IBS Mun: <vIBSMun> (NT 2024+) ou <vIBS> (schema anterior)
         If XMLDOC.selectNodes(sIBSPfx & "/gIBSMun/vIBSMun").Length > 0 Then
            Compras_XML!TotIBSMunvIBS = CDbl(Substitui(XMLDOC.selectNodes(sIBSPfx & "/gIBSMun/vIBSMun").item(0).Text, ".", ",", UM_A_UM))
         ElseIf XMLDOC.selectNodes(sIBSPfx & "/gIBSMun/vIBS").Length > 0 Then
            Compras_XML!TotIBSMunvIBS = CDbl(Substitui(XMLDOC.selectNodes(sIBSPfx & "/gIBSMun/vIBS").item(0).Text, ".", ",", UM_A_UM))
         End If
         ' Total IBS (UF + Municipal)
         Compras_XML!TotIBSvIBS = CDbl(Substitui(XMLDOC.selectNodes(sIBSPfx & "/vIBS").item(0).Text, ".", ",", UM_A_UM))
         If XMLDOC.selectNodes(sIBSPfx & "/vCredPres").Length > 0 Then
            Compras_XML!TotIBSvCredPres = CDbl(Substitui(XMLDOC.selectNodes(sIBSPfx & "/vCredPres").item(0).Text, ".", ",", UM_A_UM))
         End If
         If XMLDOC.selectNodes(sIBSPfx & "/vCredPresCondSus").Length > 0 Then
            Compras_XML!TotIBSvCredPresCondSus = CDbl(Substitui(XMLDOC.selectNodes(sIBSPfx & "/vCredPresCondSus").item(0).Text, ".", ",", UM_A_UM))
         End If
      End If
      ' Grupo CBS: <gCBS> (NT 2024+) ou <gCBSTot> (schema anterior)
      Dim sCBSPfx As String
      sCBSPfx = ""
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/gCBS").Length > 0 Then
         sCBSPfx = "nfeProc/NFe/infNFe/total/IBSCBSTot/gCBS"
      ElseIf XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/gCBSTot").Length > 0 Then
         sCBSPfx = "nfeProc/NFe/infNFe/total/IBSCBSTot/gCBSTot"
      End If
      If sCBSPfx <> "" Then
         Compras_XML!TotCBSvDif = CDbl(Substitui(XMLDOC.selectNodes(sCBSPfx & "/vDif").item(0).Text, ".", ",", UM_A_UM))
         Compras_XML!TotCBSvDevTrib = CDbl(Substitui(XMLDOC.selectNodes(sCBSPfx & "/vDevTrib").item(0).Text, ".", ",", UM_A_UM))
         Compras_XML!TotCBSvCBS = CDbl(Substitui(XMLDOC.selectNodes(sCBSPfx & "/vCBS").item(0).Text, ".", ",", UM_A_UM))
         If XMLDOC.selectNodes(sCBSPfx & "/vCredPres").Length > 0 Then
            Compras_XML!TotCBSvCredPres = CDbl(Substitui(XMLDOC.selectNodes(sCBSPfx & "/vCredPres").item(0).Text, ".", ",", UM_A_UM))
         End If
         If XMLDOC.selectNodes(sCBSPfx & "/vCredPresCondSus").Length > 0 Then
            Compras_XML!TotCBSvCredPresCondSus = CDbl(Substitui(XMLDOC.selectNodes(sCBSPfx & "/vCredPresCondSus").item(0).Text, ".", ",", UM_A_UM))
         End If
      End If
      ' vNFTot: Valor total da NF-e incluindo IBS+CBS - dentro de IBSCBSTot (schema anterior)
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/vNFTot").Length > 0 Then
         Compras_XML!vNFTot = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/vNFTot").item(0).Text, ".", ",", UM_A_UM))
      End If
   End If
   ' vNFTot: Valor total da NF-e incluindo IBS+CBS - filho direto de <total> (NT 2024+)
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/vNFTot").Length > 0 Then
      Compras_XML!vNFTot = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/vNFTot").item(0).Text, ".", ",", UM_A_UM))
   End If
   '--- ISTot - Total do Imposto Seletivo (NT 2025.002 - W31-W33) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISTot").Length > 0 Then
      Compras_XML!TotISvIS = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/ISTot/vIS").item(0).Text, ".", ",", UM_A_UM))
   End If
   '--- IBSCBSTot/gMono - Totais tributacao monofasica IBS/CBS (NT 2025.002 - W57-W59d) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/gMono").Length > 0 Then
      Compras_XML!TotIBSMonovIBSMono = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/gMono/vIBSMono").item(0).Text, ".", ",", UM_A_UM))
      Compras_XML!TotIBSMonovCBSMono = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/gMono/vCBSMono").item(0).Text, ".", ",", UM_A_UM))
      Compras_XML!TotIBSMonovIBSMonoReten = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/gMono/vIBSMonoReten").item(0).Text, ".", ",", UM_A_UM))
      Compras_XML!TotIBSMonovCBSMonoReten = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/gMono/vCBSMonoReten").item(0).Text, ".", ",", UM_A_UM))
      Compras_XML!TotIBSMonovIBSMonoRet = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/gMono/vIBSMonoRet").item(0).Text, ".", ",", UM_A_UM))
      Compras_XML!TotIBSMonovCBSMonoRet = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/total/IBSCBSTot/gMono/vCBSMonoRet").item(0).Text, ".", ",", UM_A_UM))
   End If

   '--- Modalidade do Frete (transp/modFrete - sempre presente) ---
   Compras_XML!TranspModFrete = 9
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/modFrete").Length > 0 Then
      Compras_XML!TranspModFrete = CInt(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/modFrete").item(0).Text)
   End If

   '--- Transportadora (transp/transporta - presente quando ha transportadora identificada) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta").Length > 0 Then
      'Codigo interno da transportadora cadastrada (ou recem-cadastrada) na tabela Transportadora
      Compras_XML!TranspCodigo = Codigo_transportadora
      'xNome: Razao Social ou Nome do transportador - tag <xNome>
      Compras_XML!TranspNome = UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/xNome").item(0).Text)
      'IE: Inscricao Estadual do transportador - tag <IE>
      Compras_XML!TranspInscricaoEstadual = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/IE").item(0).Text
      'xEnder: Endereco completo do transportador - tag <xEnder>
      Compras_XML!TranspEndereco = UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/xEnder").item(0).Text)
      'xMun: Municipio do transportador - tag <xMun>
      Compras_XML!TranspMunicipio = UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/xMun").item(0).Text)
      'UF: Sigla da UF do transportador - tag <UF>
      Compras_XML!TranspUF = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/UF").item(0).Text
      'CNPJ ou CPF do transportador (mutuamente exclusivos)
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/CNPJ").Length > 0 Then
         Compras_XML!TranspCNPJ_CPF = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/CNPJ").item(0).Text
      ElseIf XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/CPF").Length > 0 Then
         Compras_XML!TranspCNPJ_CPF = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/transporta/CPF").item(0).Text
      End If
   End If

   '--- Retencao ICMS do Transporte (transp/retTransp - presente quando ha retencao de ICMS pelo transportador) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/retTransp").Length > 0 Then
      'vServ: Valor do Servico de Transporte - tag <vServ>
      Compras_XML!TranspValorServico = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/retTransp/vServ").item(0).Text, ".", ",", UM_A_UM))
      'vBCRet: BC da Retencao do ICMS Transporte - tag <vBCRet>
      Compras_XML!TranspBCRetencaoICMS = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/retTransp/vBCRet").item(0).Text, ".", ",", UM_A_UM))
      'pICMSRet: Aliquota da Retencao do ICMS Transporte - tag <pICMSRet>
      Compras_XML!TranspAliquotaRetencaoICMS = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/retTransp/pICMSRet").item(0).Text, ".", ",", UM_A_UM))
      'vICMSRet: Valor da Retencao do ICMS Transporte - tag <vICMSRet>
      Compras_XML!TranspValorRetencaoICMS = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/retTransp/vICMSRet").item(0).Text, ".", ",", UM_A_UM))
      'CFOP: CFOP da prestacao do servico de transporte - tag <CFOP>
      Compras_XML!TranspCFOP = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/retTransp/CFOP").item(0).Text
      'cMunFG: Municipio de ocorrencia do fato gerador do ICMS do transporte - tag <cMunFG>
      Compras_XML!TranspcMunFG = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/retTransp/cMunFG").item(0).Text
   End If

   '--- Veiculo Transportador (transp/veicTransp - opcional) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/veicTransp").Length > 0 Then
      'placa: Placa do veiculo - tag <placa>
      Compras_XML!TranspPlaca = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/veicTransp/placa").item(0).Text
      'UF: UF de licenciamento do veiculo - tag <UF>
      Compras_XML!TranspPlacaUF = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/veicTransp/UF").item(0).Text
      'RNTC: Registro Nacional de Transportador de Carga - tag <RNTC> (opcional)
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/veicTransp/RNTC").Length > 0 Then
         Compras_XML!TranspRNTC = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/veicTransp/RNTC").item(0).Text
      End If
   End If

   '--- Reboque (transp/reboque - opcional) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/reboque").Length > 0 Then
      Compras_XML!TranspReboquePlaca = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/reboque/placa").item(0).Text
      Compras_XML!TranspReboquePlacaUF = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/reboque/UF").item(0).Text
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/reboque/RNTC").Length > 0 Then
         Compras_XML!TranspReboqueRNTC = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/reboque/RNTC").item(0).Text
      End If
   End If

   '--- Volumes (transp/vol - opcional, pode haver multiplos; registra apenas o primeiro) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol").Length > 0 Then
      'qVol: Quantidade de Volumes - tag <qVol> (obrigatoria)
      Compras_XML!VolumeQuantidade = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/qVol").item(0).Text
      'esp: Especie dos Volumes - tag <esp> (obrigatoria)
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/esp").Length > 0 Then
         Compras_XML!VolumeEspecie = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/esp").item(0).Text
      End If
      'marca: Marca dos Volumes - tag <marca> (opcional)
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/marca").Length > 0 Then
         Compras_XML!VolumeMarca = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/marca").item(0).Text
      End If
      'nVol: Numeracao dos Volumes - tag <nVol> (opcional)
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/nVol").Length > 0 Then
         Compras_XML!VolumeNumeracao = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/nVol").item(0).Text
      End If
      'pesoL: Peso Liquido em kg - tag <pesoL> (opcional)
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/pesoL").Length > 0 Then
         Compras_XML!VolumePesoLiquido = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/pesoL").item(0).Text, ".", ",", UM_A_UM))
      End If
      'pesoB: Peso Bruto em kg - tag <pesoB> (opcional)
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/pesoB").Length > 0 Then
         Compras_XML!VolumePesoBruto = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/pesoB").item(0).Text, ".", ",", UM_A_UM))
      End If
      'nLacre: Numero dos lacres dos volumes - tag <nLacre> (opcional)
      If XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/lacres/nLacre").Length > 0 Then
         Compras_XML!lacres = XMLDOC.selectNodes("nfeProc/NFe/infNFe/transp/vol/lacres/nLacre").item(0).Text
      End If
   End If

   '--- Fatura (cobr/fat - opcional) ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/fat").Length > 0 Then
      'nFat: Numero da fatura - tag <nFat>
      Compras_XML!NumeroFatura = XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/fat/nFat").item(0).Text
      'vOrig: Valor original da fatura - tag <vOrig>
      Compras_XML!ValorOriginalFatura = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/fat/vOrig").item(0).Text, ".", ",", UM_A_UM))
      'vDesc: Valor do desconto da fatura - tag <vDesc>
      Compras_XML!ValorDescontoFatura = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/fat/vDesc").item(0).Text, ".", ",", UM_A_UM))
      'vLiq: Valor liquido da fatura (vOrig - vDesc) - tag <vLiq>
      Compras_XML!ValorLiquidoFatura = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/fat/vLiq").item(0).Text, ".", ",", UM_A_UM))
   End If

   '--- Primeira duplicata registrada na entrada; demais parcelas vao para ContasPagar ---
   If XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").Length > 0 Then
      'nDup: Numero da duplicata - tag <nDup>
      Compras_XML!NumeroDuplicata = XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").item(0).selectNodes("nDup").item(0).Text
      'dVenc: Data de vencimento da duplicata - tag <dVenc>
      Compras_XML!VencimentoDuplicata = XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").item(0).selectNodes("dVenc").item(0).Text
      'vDup: Valor da duplicata - tag <vDup>
      Compras_XML!ValorDuplicata = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").item(0).selectNodes("vDup").item(0).Text, ".", ",", UM_A_UM))
   End If

   Compras_XML.Update
  '------------- FIM DO CADASTRO DA ENTRADA ------------------
      
      
      
       '-------- IMPORTANDO OS ITENS DA NOTA ---------
       'Contando quantos itens tem o na det (detalhes)
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
        
           '==============================================================
           ' GRUPO ICMS - Imposto sobre Circulacao de Mercadorias
           ' Detecta o subgrupo presente no XML e le as tags conforme
           ' a legislacao de cada CST / CSOSN aplicavel
           '==============================================================

           'Zera todas as variaveis ICMS antes de cada item
           subgrupoICMS = ""
           csticms = ""
           ICMSBase = 0: ICMSAliquota = 0: ICMSValor = 0
           ICMSmodBC = 0: ICMSpRedBC = 0: ICMSvICMSOp = 0: ICMSpDif = 0
           ICMSvICMSDeson = 0: ICMSmotDesICMS = ""
           ICMSmodBCST = 0: ICMSpRedBCST = 0: ICMSpMVAST = 0
           ICMSvBCST = 0: ICMSpICMSST = 0: ICMSvICMSST = 0
           ICMSvBCFCPST = 0: ICMSpFCPST = 0: ICMSvFCPST = 0
           ICMSvBCSTRet = 0: ICMSpST = 0: ICMSvICMSSubstituto = 0
           ICMSvICMSSTRet = 0: ICMSvBCFCPSTRet = 0: ICMSpFCPSTRet = 0: ICMSvFCPSTRet = 0

           'Detecta qual subgrupo ICMS existe no item atual
           Dim oICMS As Object
           Set oICMS = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/ICMS").item(0)

           If Not oICMS Is Nothing Then
              If oICMS.selectNodes("ICMS00").Length > 0 Then subgrupoICMS = "ICMS00"
              If oICMS.selectNodes("ICMS10").Length > 0 Then subgrupoICMS = "ICMS10"
              If oICMS.selectNodes("ICMS20").Length > 0 Then subgrupoICMS = "ICMS20"
              If oICMS.selectNodes("ICMS40").Length > 0 Then subgrupoICMS = "ICMS40"
              If oICMS.selectNodes("ICMS41").Length > 0 Then subgrupoICMS = "ICMS41"
              If oICMS.selectNodes("ICMS51").Length > 0 Then subgrupoICMS = "ICMS51"
              If oICMS.selectNodes("ICMS60").Length > 0 Then subgrupoICMS = "ICMS60"
              If oICMS.selectNodes("ICMS70").Length > 0 Then subgrupoICMS = "ICMS70"
              If oICMS.selectNodes("ICMS90").Length > 0 Then subgrupoICMS = "ICMS90"
              'Simples Nacional: CSOSN no lugar de CST
              If oICMS.selectNodes("ICMSSN101").Length > 0 Then subgrupoICMS = "ICMSSN101"
              If oICMS.selectNodes("ICMSSN102").Length > 0 Then subgrupoICMS = "ICMSSN102"
              If oICMS.selectNodes("ICMSSN201").Length > 0 Then subgrupoICMS = "ICMSSN201"
              If oICMS.selectNodes("ICMSSN202").Length > 0 Then subgrupoICMS = "ICMSSN202"
              If oICMS.selectNodes("ICMSSN500").Length > 0 Then subgrupoICMS = "ICMSSN500"
              If oICMS.selectNodes("ICMSSN900").Length > 0 Then subgrupoICMS = "ICMSSN900"
           End If

           Select Case subgrupoICMS

              Case "ICMS00"
                 'Tributacao plena - CST 00
                 'orig:  Origem da mercadoria (0=Nacional, 1=Estrangeira, etc.)
                 'CST:   Codigo de Situacao Tributaria do ICMS
                 'modBC: Modalidade de determinacao da base de calculo
                 'vBC:   Valor da base de calculo do ICMS
                 'pICMS: Aliquota do ICMS (%)
                 'vICMS: Valor do ICMS
                 csticms = oICMS.selectNodes("ICMS00/orig").item(0).Text & oICMS.selectNodes("ICMS00/CST").item(0).Text
                 ICMSmodBC = CInt(oICMS.selectNodes("ICMS00/modBC").item(0).Text)
                 ICMSBase = CDbl(Substitui(oICMS.selectNodes("ICMS00/vBC").item(0).Text, ".", ",", UM_A_UM))
                 ICMSAliquota = CDbl(Substitui(oICMS.selectNodes("ICMS00/pICMS").item(0).Text, ".", ",", UM_A_UM))
                 ICMSValor = CDbl(Substitui(oICMS.selectNodes("ICMS00/vICMS").item(0).Text, ".", ",", UM_A_UM))

              Case "ICMS10"
                 'Tributacao com ST - CST 10
                 'orig:     Origem da mercadoria
                 'CST:      Codigo de Situacao Tributaria
                 'modBC:    Modalidade da base de calculo do ICMS proprio
                 'vBC:      Base de calculo do ICMS proprio
                 'pICMS:    Aliquota do ICMS proprio
                 'vICMS:    Valor do ICMS proprio
                 'modBCST:  Modalidade da base de calculo do ICMS ST
                 'pRedBCST: Percentual de reducao da base de calculo ST
                 'vBCST:    Base de calculo do ICMS ST
                 'pICMSST:  Aliquota do ICMS ST
                 'vICMSST:  Valor do ICMS ST
                 'vBCFCPST: Base de calculo do FCP retido por ST
                 'pFCPST:   Percentual do FCP retido por ST
                 'vFCPST:   Valor do FCP retido por ST
                 csticms = oICMS.selectNodes("ICMS10/orig").item(0).Text & oICMS.selectNodes("ICMS10/CST").item(0).Text
                 ICMSmodBC = CInt(oICMS.selectNodes("ICMS10/modBC").item(0).Text)
                 ICMSBase = CDbl(Substitui(oICMS.selectNodes("ICMS10/vBC").item(0).Text, ".", ",", UM_A_UM))
                 ICMSAliquota = CDbl(Substitui(oICMS.selectNodes("ICMS10/pICMS").item(0).Text, ".", ",", UM_A_UM))
                 ICMSValor = CDbl(Substitui(oICMS.selectNodes("ICMS10/vICMS").item(0).Text, ".", ",", UM_A_UM))
                 ICMSmodBCST = CInt(oICMS.selectNodes("ICMS10/modBCST").item(0).Text)
                 ICMSpRedBCST = CDbl(Substitui(oICMS.selectNodes("ICMS10/pRedBCST").item(0).Text, ".", ",", UM_A_UM))
                 ICMSvBCST = CDbl(Substitui(oICMS.selectNodes("ICMS10/vBCST").item(0).Text, ".", ",", UM_A_UM))
                 ICMSpICMSST = CDbl(Substitui(oICMS.selectNodes("ICMS10/pICMSST").item(0).Text, ".", ",", UM_A_UM))
                 ICMSvICMSST = CDbl(Substitui(oICMS.selectNodes("ICMS10/vICMSST").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS10/vBCFCPST").Length > 0 Then ICMSvBCFCPST = CDbl(Substitui(oICMS.selectNodes("ICMS10/vBCFCPST").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS10/pFCPST").Length > 0 Then ICMSpFCPST = CDbl(Substitui(oICMS.selectNodes("ICMS10/pFCPST").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS10/vFCPST").Length > 0 Then ICMSvFCPST = CDbl(Substitui(oICMS.selectNodes("ICMS10/vFCPST").item(0).Text, ".", ",", UM_A_UM))

              Case "ICMS20"
                 'Reducao de base de calculo - CST 20
                 'orig:   Origem da mercadoria
                 'CST:    Codigo de Situacao Tributaria
                 'modBC:  Modalidade da base de calculo
                 'pRedBC: Percentual de reducao da base de calculo
                 'vBC:    Base de calculo reduzida do ICMS
                 'pICMS:  Aliquota do ICMS
                 'vICMS:  Valor do ICMS
                 csticms = oICMS.selectNodes("ICMS20/orig").item(0).Text & oICMS.selectNodes("ICMS20/CST").item(0).Text
                 ICMSmodBC = CInt(oICMS.selectNodes("ICMS20/modBC").item(0).Text)
                 ICMSpRedBC = CDbl(Substitui(oICMS.selectNodes("ICMS20/pRedBC").item(0).Text, ".", ",", UM_A_UM))
                 ICMSBase = CDbl(Substitui(oICMS.selectNodes("ICMS20/vBC").item(0).Text, ".", ",", UM_A_UM))
                 ICMSAliquota = CDbl(Substitui(oICMS.selectNodes("ICMS20/pICMS").item(0).Text, ".", ",", UM_A_UM))
                 ICMSValor = CDbl(Substitui(oICMS.selectNodes("ICMS20/vICMS").item(0).Text, ".", ",", UM_A_UM))

              Case "ICMS40"
                 'Isencao, nao-tributado ou suspensao - CST 40
                 'orig:       Origem da mercadoria
                 'CST:        Codigo de Situacao Tributaria (40=Isento, 41=Nao-Tributado, 50=Suspensao)
                 'vICMSDeson: Valor do ICMS desonerado (opcional)
                 'motDesICMS: Motivo da desoneracao (opcional - codigo numerico)
                 csticms = oICMS.selectNodes("ICMS40/orig").item(0).Text & oICMS.selectNodes("ICMS40/CST").item(0).Text
                 If oICMS.selectNodes("ICMS40/vICMSDeson").Length > 0 Then ICMSvICMSDeson = CDbl(Substitui(oICMS.selectNodes("ICMS40/vICMSDeson").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS40/motDesICMS").Length > 0 Then ICMSmotDesICMS = oICMS.selectNodes("ICMS40/motDesICMS").item(0).Text

              Case "ICMS41"
                 'Nao-tributado - CST 41
                 'orig:       Origem da mercadoria
                 'CST:        Codigo de Situacao Tributaria
                 'vICMSDeson: Valor do ICMS desonerado (opcional)
                 'motDesICMS: Motivo da desoneracao (opcional)
                 csticms = oICMS.selectNodes("ICMS41/orig").item(0).Text & oICMS.selectNodes("ICMS41/CST").item(0).Text
                 If oICMS.selectNodes("ICMS41/vICMSDeson").Length > 0 Then ICMSvICMSDeson = CDbl(Substitui(oICMS.selectNodes("ICMS41/vICMSDeson").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS41/motDesICMS").Length > 0 Then ICMSmotDesICMS = oICMS.selectNodes("ICMS41/motDesICMS").item(0).Text

              Case "ICMS51"
                 'Diferimento - CST 51
                 'orig:    Origem da mercadoria
                 'CST:     Codigo de Situacao Tributaria
                 'modBC:   Modalidade da base de calculo
                 'pRedBC:  Percentual de reducao da base de calculo
                 'vBC:     Base de calculo do ICMS
                 'pICMS:   Aliquota nominal do ICMS
                 'vICMSOp: Valor do ICMS da operacao (antes do diferimento)
                 'pDif:    Percentual do diferimento
                 'vICMS:   Valor do ICMS efetivamente devido (apos diferimento)
                 csticms = oICMS.selectNodes("ICMS51/orig").item(0).Text & oICMS.selectNodes("ICMS51/CST").item(0).Text
                 ICMSmodBC = CInt(oICMS.selectNodes("ICMS51/modBC").item(0).Text)
                 ICMSpRedBC = CDbl(Substitui(oICMS.selectNodes("ICMS51/pRedBC").item(0).Text, ".", ",", UM_A_UM))
                 ICMSBase = CDbl(Substitui(oICMS.selectNodes("ICMS51/vBC").item(0).Text, ".", ",", UM_A_UM))
                 ICMSAliquota = CDbl(Substitui(oICMS.selectNodes("ICMS51/pICMS").item(0).Text, ".", ",", UM_A_UM))
                 ICMSvICMSOp = CDbl(Substitui(oICMS.selectNodes("ICMS51/vICMSOp").item(0).Text, ".", ",", UM_A_UM))
                 ICMSpDif = CDbl(Substitui(oICMS.selectNodes("ICMS51/pDif").item(0).Text, ".", ",", UM_A_UM))
                 ICMSValor = CDbl(Substitui(oICMS.selectNodes("ICMS51/vICMS").item(0).Text, ".", ",", UM_A_UM))

              Case "ICMS60"
                 'ICMS cobrado por ST retido anteriormente - CST 60
                 'orig:           Origem da mercadoria
                 'CST:            Codigo de Situacao Tributaria
                 'vBCSTRet:       Base de calculo do ICMS ST retido anteriormente (opcional - ausente em CST 60 simplificado)
                 'pST:            Aliquota suportada pelo consumidor final (opcional)
                 'vICMSSubstituto:Valor do ICMS do substituto tributario (opcional)
                 'vICMSSTRet:     Valor do ICMS ST retido anteriormente (opcional)
                 'vBCFCPSTRet:    Base de calculo do FCP ST retido (opcional)
                 'pFCPSTRet:      Percentual do FCP ST retido (opcional)
                 'vFCPSTRet:      Valor do FCP ST retido (opcional)
                 csticms = oICMS.selectNodes("ICMS60/orig").item(0).Text & oICMS.selectNodes("ICMS60/CST").item(0).Text
                 If oICMS.selectNodes("ICMS60/vBCSTRet").Length > 0 Then
                    ICMSvBCSTRet = CDbl(Substitui(oICMS.selectNodes("ICMS60/vBCSTRet").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oICMS.selectNodes("ICMS60/pST").Length > 0 Then
                    ICMSpST = CDbl(Substitui(oICMS.selectNodes("ICMS60/pST").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oICMS.selectNodes("ICMS60/vICMSSubstituto").Length > 0 Then
                    ICMSvICMSSubstituto = CDbl(Substitui(oICMS.selectNodes("ICMS60/vICMSSubstituto").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oICMS.selectNodes("ICMS60/vICMSSTRet").Length > 0 Then
                    ICMSvICMSSTRet = CDbl(Substitui(oICMS.selectNodes("ICMS60/vICMSSTRet").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oICMS.selectNodes("ICMS60/vBCFCPSTRet").Length > 0 Then
                    ICMSvBCFCPSTRet = CDbl(Substitui(oICMS.selectNodes("ICMS60/vBCFCPSTRet").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oICMS.selectNodes("ICMS60/pFCPSTRet").Length > 0 Then
                    ICMSpFCPSTRet = CDbl(Substitui(oICMS.selectNodes("ICMS60/pFCPSTRet").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oICMS.selectNodes("ICMS60/vFCPSTRet").Length > 0 Then
                    ICMSvFCPSTRet = CDbl(Substitui(oICMS.selectNodes("ICMS60/vFCPSTRet").item(0).Text, ".", ",", UM_A_UM))
                 End If

              Case "ICMS70"
                 'Reducao de BC e cobranca de ICMS por ST - CST 70
                 'orig:    Origem da mercadoria
                 'CST:     Codigo de Situacao Tributaria
                 'modBC:   Modalidade da base de calculo do ICMS proprio
                 'pRedBC:  Percentual de reducao da base de calculo
                 'vBC:     Base de calculo do ICMS proprio
                 'pICMS:   Aliquota do ICMS proprio
                 'vICMS:   Valor do ICMS proprio
                 'modBCST: Modalidade da base de calculo do ICMS ST
                 'pMVAST:  Percentual de Margem de Valor Adicionado ST
                 'vBCST:   Base de calculo do ICMS ST
                 'pICMSST: Aliquota do ICMS ST
                 'vICMSST: Valor do ICMS ST
                 'vBCFCPST/pFCPST/vFCPST: FCP ST (opcional)
                 csticms = oICMS.selectNodes("ICMS70/orig").item(0).Text & oICMS.selectNodes("ICMS70/CST").item(0).Text
                 ICMSmodBC = CInt(oICMS.selectNodes("ICMS70/modBC").item(0).Text)
                 ICMSpRedBC = CDbl(Substitui(oICMS.selectNodes("ICMS70/pRedBC").item(0).Text, ".", ",", UM_A_UM))
                 ICMSBase = CDbl(Substitui(oICMS.selectNodes("ICMS70/vBC").item(0).Text, ".", ",", UM_A_UM))
                 ICMSAliquota = CDbl(Substitui(oICMS.selectNodes("ICMS70/pICMS").item(0).Text, ".", ",", UM_A_UM))
                 ICMSValor = CDbl(Substitui(oICMS.selectNodes("ICMS70/vICMS").item(0).Text, ".", ",", UM_A_UM))
                 ICMSmodBCST = CInt(oICMS.selectNodes("ICMS70/modBCST").item(0).Text)
                 ICMSpMVAST = CDbl(Substitui(oICMS.selectNodes("ICMS70/pMVAST").item(0).Text, ".", ",", UM_A_UM))
                 ICMSvBCST = CDbl(Substitui(oICMS.selectNodes("ICMS70/vBCST").item(0).Text, ".", ",", UM_A_UM))
                 ICMSpICMSST = CDbl(Substitui(oICMS.selectNodes("ICMS70/pICMSST").item(0).Text, ".", ",", UM_A_UM))
                 ICMSvICMSST = CDbl(Substitui(oICMS.selectNodes("ICMS70/vICMSST").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS70/vBCFCPST").Length > 0 Then ICMSvBCFCPST = CDbl(Substitui(oICMS.selectNodes("ICMS70/vBCFCPST").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS70/pFCPST").Length > 0 Then ICMSpFCPST = CDbl(Substitui(oICMS.selectNodes("ICMS70/pFCPST").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS70/vFCPST").Length > 0 Then ICMSvFCPST = CDbl(Substitui(oICMS.selectNodes("ICMS70/vFCPST").item(0).Text, ".", ",", UM_A_UM))

              Case "ICMS90"
                 'Outros - CST 90 (regime normal, situacoes nao enquadradas)
                 'orig:    Origem da mercadoria
                 'CST:     Codigo de Situacao Tributaria
                 'modBC:   Modalidade da base de calculo
                 'vBC:     Base de calculo do ICMS
                 'pICMS:   Aliquota do ICMS
                 'vICMS:   Valor do ICMS
                 'pRedBC/pMVAST/vBCST/vICMSST/vBCFCPST/pFCPST/vFCPST: opcionais (ST se houver)
                 csticms = oICMS.selectNodes("ICMS90/orig").item(0).Text & oICMS.selectNodes("ICMS90/CST").item(0).Text
                 ICMSmodBC = CInt(oICMS.selectNodes("ICMS90/modBC").item(0).Text)
                 ICMSBase = CDbl(Substitui(oICMS.selectNodes("ICMS90/vBC").item(0).Text, ".", ",", UM_A_UM))
                 ICMSAliquota = CDbl(Substitui(oICMS.selectNodes("ICMS90/pICMS").item(0).Text, ".", ",", UM_A_UM))
                 ICMSValor = CDbl(Substitui(oICMS.selectNodes("ICMS90/vICMS").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS90/pRedBC").Length > 0 Then ICMSpRedBC = CDbl(Substitui(oICMS.selectNodes("ICMS90/pRedBC").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS90/pMVAST").Length > 0 Then ICMSpMVAST = CDbl(Substitui(oICMS.selectNodes("ICMS90/pMVAST").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS90/vBCST").Length > 0 Then ICMSvBCST = CDbl(Substitui(oICMS.selectNodes("ICMS90/vBCST").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS90/vICMSST").Length > 0 Then ICMSvICMSST = CDbl(Substitui(oICMS.selectNodes("ICMS90/vICMSST").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS90/vBCFCPST").Length > 0 Then ICMSvBCFCPST = CDbl(Substitui(oICMS.selectNodes("ICMS90/vBCFCPST").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS90/pFCPST").Length > 0 Then ICMSpFCPST = CDbl(Substitui(oICMS.selectNodes("ICMS90/pFCPST").item(0).Text, ".", ",", UM_A_UM))
                 If oICMS.selectNodes("ICMS90/vFCPST").Length > 0 Then ICMSvFCPST = CDbl(Substitui(oICMS.selectNodes("ICMS90/vFCPST").item(0).Text, ".", ",", UM_A_UM))

              Case "ICMSSN101", "ICMSSN102", "ICMSSN201", "ICMSSN202", "ICMSSN500", "ICMSSN900"
                 'Simples Nacional - usa CSOSN no lugar de CST
                 'orig:  Origem da mercadoria
                 'CSOSN: Codigo de Situacao da Operacao no Simples Nacional
                 Dim sSN As String: sSN = subgrupoICMS
                 csticms = oICMS.selectNodes(sSN & "/orig").item(0).Text & oICMS.selectNodes(sSN & "/CSOSN").item(0).Text

           End Select

           '==============================================================
           ' GRUPO IPI - Imposto sobre Produtos Industrializados
           '==============================================================

           'Zera variaveis IPI
           IPIcEnq = "": cstipi = "": IPIvBC = 0: IPIpIPI = 0: IPIvIPI = 0

           Dim oIPI As Object
           Set oIPI = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/IPI").item(0)

           If Not oIPI Is Nothing Then
              'cEnq: Codigo de enquadramento legal do IPI (ex: 999 para nao tributados)
              IPIcEnq = oIPI.selectNodes("cEnq").item(0).Text

              If oIPI.selectNodes("IPITrib").Length > 0 Then
                 'Subgrupo IPITrib: produto tributado pelo IPI
                 'CST:  Codigo de Situacao Tributaria do IPI (50=entrada tributada)
                 'vBC:  Base de calculo do IPI
                 'pIPI: Aliquota do IPI (%)
                 'vIPI: Valor do IPI
                 cstipi = oIPI.selectNodes("IPITrib/CST").item(0).Text
                 IPIvBC = CDbl(Substitui(oIPI.selectNodes("IPITrib/vBC").item(0).Text, ".", ",", UM_A_UM))
                 IPIpIPI = CDbl(Substitui(oIPI.selectNodes("IPITrib/pIPI").item(0).Text, ".", ",", UM_A_UM))
                 IPIvIPI = CDbl(Substitui(oIPI.selectNodes("IPITrib/vIPI").item(0).Text, ".", ",", UM_A_UM))
              ElseIf oIPI.selectNodes("IPINT").Length > 0 Then
                 'Subgrupo IPINT: produto nao tributado pelo IPI
                 'CST:  Codigo de Situacao Tributaria (01=Entrada tributada com aliq zero, 02=Entrada isenta, etc.)
                 'vBC/pIPI/vIPI: zerados pois nao ha tributacao
                 cstipi = oIPI.selectNodes("IPINT/CST").item(0).Text
                 IPIvBC = 0
                 IPIpIPI = 0
                 IPIvIPI = 0
              End If
           End If

           '==============================================================
           ' GRUPO PIS - Programa de Integracao Social
           '==============================================================

           'Zera variaveis PIS
           pisCST = "": PISvBC = 0: PISpPIS = 0: PISvPIS = 0
           PISqBCProd = 0: PISvAliqProd = 0

           Dim oPIS As Object
           Set oPIS = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/PIS").item(0)

           If Not oPIS Is Nothing Then

              If oPIS.selectNodes("PISAliq").Length > 0 Then
                 'Subgrupo PISAliq: CST 01 ou 02 - PIS com aliquota em percentual
                 'CST:  01=Operacao tributavel BC=valor operacao, 02=Operacao tributavel BC=valor operacao (aliq diferenciada)
                 'vBC:  Base de calculo do PIS
                 'pPIS: Aliquota do PIS (%)
                 'vPIS: Valor do PIS
                 pisCST = oPIS.selectNodes("PISAliq/CST").item(0).Text
                 PISvBC = CDbl(Substitui(oPIS.selectNodes("PISAliq/vBC").item(0).Text, ".", ",", UM_A_UM))
                 PISpPIS = CDbl(Substitui(oPIS.selectNodes("PISAliq/pPIS").item(0).Text, ".", ",", UM_A_UM))
                 PISvPIS = CDbl(Substitui(oPIS.selectNodes("PISAliq/vPIS").item(0).Text, ".", ",", UM_A_UM))

              ElseIf oPIS.selectNodes("PISQtde").Length > 0 Then
                 'Subgrupo PISQtde: CST 03 - PIS com aliquota por unidade de produto
                 'CST:       03=Operacao tributavel por unidade de medida de produto
                 'qBCProd:   Quantidade vendida (unidades de medida)
                 'vAliqProd: Aliquota do PIS em reais por unidade de produto
                 'vPIS:      Valor do PIS (qBCProd x vAliqProd)
                 pisCST = oPIS.selectNodes("PISQtde/CST").item(0).Text
                 PISqBCProd = CDbl(Substitui(oPIS.selectNodes("PISQtde/qBCProd").item(0).Text, ".", ",", UM_A_UM))
                 PISvAliqProd = CDbl(Substitui(oPIS.selectNodes("PISQtde/vAliqProd").item(0).Text, ".", ",", UM_A_UM))
                 PISvPIS = CDbl(Substitui(oPIS.selectNodes("PISQtde/vPIS").item(0).Text, ".", ",", UM_A_UM))

              ElseIf oPIS.selectNodes("PISNT").Length > 0 Then
                 'Subgrupo PISNT: CST 04 a 09 - PIS nao tributado
                 'CST: 04=Op. isenta, 05=Op. imune, 06=Op. c/ suspensao, 07=Op. sem incidencia,
                 '     08=Op. sem incidencia (exportacao), 09=Op. sem incidencia (outros)
                 'vBC/pPIS/vPIS: zerados pois nao ha tributacao
                 pisCST = oPIS.selectNodes("PISNT/CST").item(0).Text

              ElseIf oPIS.selectNodes("PISOutr").Length > 0 Then
                 'Subgrupo PISOutr: CST 99 - outras operacoes
                 'CST:  99=Outras operacoes de saida
                 'vBC:  Base de calculo do PIS
                 'pPIS: Aliquota do PIS (%)
                 'vPIS: Valor do PIS
                 pisCST = oPIS.selectNodes("PISOutr/CST").item(0).Text
                 PISvBC = CDbl(Substitui(oPIS.selectNodes("PISOutr/vBC").item(0).Text, ".", ",", UM_A_UM))
                 PISpPIS = CDbl(Substitui(oPIS.selectNodes("PISOutr/pPIS").item(0).Text, ".", ",", UM_A_UM))
                 PISvPIS = CDbl(Substitui(oPIS.selectNodes("PISOutr/vPIS").item(0).Text, ".", ",", UM_A_UM))
              End If
           End If

           '==============================================================
           ' GRUPO COFINS - Contribuicao para Financiamento da Seguridade Social
           '==============================================================

           'Zera variaveis COFINS
           cofinsCST = "": COFINSvBC = 0: COFINSpCOFINS = 0: COFINSvCOFINS = 0
           COFINSqBCProd = 0: COFINSvAliqProd = 0

           Dim oCOFINS As Object
           Set oCOFINS = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/COFINS").item(0)

           If Not oCOFINS Is Nothing Then

              If oCOFINS.selectNodes("COFINSAliq").Length > 0 Then
                 'Subgrupo COFINSAliq: CST 01 ou 02 - COFINS com aliquota em percentual
                 'CST:     01=Op. tributavel BC=valor operacao, 02=Op. tributavel BC=valor operacao (aliq diferenciada)
                 'vBC:     Base de calculo do COFINS
                 'pCOFINS: Aliquota do COFINS (%)
                 'vCOFINS: Valor do COFINS
                 cofinsCST = oCOFINS.selectNodes("COFINSAliq/CST").item(0).Text
                 COFINSvBC = CDbl(Substitui(oCOFINS.selectNodes("COFINSAliq/vBC").item(0).Text, ".", ",", UM_A_UM))
                 COFINSpCOFINS = CDbl(Substitui(oCOFINS.selectNodes("COFINSAliq/pCOFINS").item(0).Text, ".", ",", UM_A_UM))
                 COFINSvCOFINS = CDbl(Substitui(oCOFINS.selectNodes("COFINSAliq/vCOFINS").item(0).Text, ".", ",", UM_A_UM))

              ElseIf oCOFINS.selectNodes("COFINSQtde").Length > 0 Then
                 'Subgrupo COFINSQtde: CST 03 - COFINS com aliquota por unidade de produto
                 'CST:       03=Operacao tributavel por unidade de medida de produto
                 'qBCProd:   Quantidade vendida (unidades de medida)
                 'vAliqProd: Aliquota do COFINS em reais por unidade de produto
                 'vCOFINS:   Valor do COFINS (qBCProd x vAliqProd)
                 cofinsCST = oCOFINS.selectNodes("COFINSQtde/CST").item(0).Text
                 COFINSqBCProd = CDbl(Substitui(oCOFINS.selectNodes("COFINSQtde/qBCProd").item(0).Text, ".", ",", UM_A_UM))
                 COFINSvAliqProd = CDbl(Substitui(oCOFINS.selectNodes("COFINSQtde/vAliqProd").item(0).Text, ".", ",", UM_A_UM))
                 COFINSvCOFINS = CDbl(Substitui(oCOFINS.selectNodes("COFINSQtde/vCOFINS").item(0).Text, ".", ",", UM_A_UM))

              ElseIf oCOFINS.selectNodes("COFINSNT").Length > 0 Then
                 'Subgrupo COFINSNT: CST 04 a 09 - COFINS nao tributado
                 'CST: 04=Op. isenta, 05=Op. imune, 06=Op. c/ suspensao, 07=Op. sem incidencia,
                 '     08=Op. sem incidencia (exportacao), 09=Op. sem incidencia (outros)
                 'vBC/pCOFINS/vCOFINS: zerados pois nao ha tributacao
                 cofinsCST = oCOFINS.selectNodes("COFINSNT/CST").item(0).Text

              ElseIf oCOFINS.selectNodes("COFINSOutr").Length > 0 Then
                 'Subgrupo COFINSOutr: CST 99 - outras operacoes
                 'CST:     99=Outras operacoes de saida
                 'vBC:     Base de calculo do COFINS
                 'pCOFINS: Aliquota do COFINS (%)
                 'vCOFINS: Valor do COFINS
                 cofinsCST = oCOFINS.selectNodes("COFINSOutr/CST").item(0).Text
                 COFINSvBC = CDbl(Substitui(oCOFINS.selectNodes("COFINSOutr/vBC").item(0).Text, ".", ",", UM_A_UM))
                 COFINSpCOFINS = CDbl(Substitui(oCOFINS.selectNodes("COFINSOutr/pCOFINS").item(0).Text, ".", ",", UM_A_UM))
                 COFINSvCOFINS = CDbl(Substitui(oCOFINS.selectNodes("COFINSOutr/vCOFINS").item(0).Text, ".", ",", UM_A_UM))
              End If
           End If
           
          '==============================================================
          ' IDENTIFICACAO DO PRODUTO INTERNO
          ' Ordem de busca:
          '   1. VinculoXMLProduto por IDFornecedor + cProd  (mais confiavel - embalagem mapeada)
          '   2. VinculoXMLProduto por EANEmbalagem           (fallback por EAN da caixa)
          '   3. Produtos.EAN por cEAN                        (EAN da embalagem = EAN interno)
          '   4. Produtos.EAN por cEANTrib                    (EAN da unidade tributavel = EAN interno)
          '==============================================================
          cEAN = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/cEAN").item(0).Text
          xCProd = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/cProd").item(0).Text
          xVUnCom = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/vUnCom").item(0).Text, ".", ",", UM_A_UM))

          Dim cEANTrib As String
          If XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/cEANTrib").Length > 0 Then
             cEANTrib = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/cEANTrib").item(0).Text
          Else
             cEANTrib = ""
          End If

          'qCom e qTrib: usados para calcular fracionamento quando cEAN(caixa) != cEANTrib(unidade)
          Dim xQCom  As Double
          Dim xQTrib As Double
          xQCom = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/qCom").item(0).Text, ".", ",", UM_A_UM))
          If XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/qTrib").Length > 0 Then
             xQTrib = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/qTrib").item(0).Text, ".", ",", UM_A_UM))
          Else
             xQTrib = xQCom
          End If

          qtfracao = 1
          CodProduto = 0
          Dim bVinculoBusca1 As Boolean
          bVinculoBusca1 = False
          Dim sVinNCM As String, sVinCEST As String, sVinCFOP As String
          Dim oDetVin As Object
          sVinNCM = ""
          sVinCEST = ""
          sVinCFOP = "0"
          Set oDetVin = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i)
          If oDetVin.selectNodes("prod/NCM").Length > 0 Then sVinNCM = oDetVin.selectNodes("prod/NCM").item(0).Text
          If oDetVin.selectNodes("prod/CEST").Length > 0 Then sVinCEST = oDetVin.selectNodes("prod/CEST").item(0).Text
          If oDetVin.selectNodes("prod/CFOP").Length > 0 Then sVinCFOP = oDetVin.selectNodes("prod/CFOP").item(0).Text
          Set oDetVin = Nothing

          '--- Busca 1: VinculoXMLProduto por IDFornecedor + cProd ---
          RsOpen TbVinculo, "SELECT IDProduto, Fracionamento FROM VinculoXMLProduto " & _
                            " WHERE IDFornecedor = " & Codigo_fornecedor & _
                            "   AND cProd = '" & Replace(xCProd, "'" , "''") & "'"
          If TbVinculo.RecordCount > 0 Then
             CodProduto = TbVinculo!IDProduto
             qtfracao = IIf(TbVinculo!Fracionamento > 0, TbVinculo!Fracionamento, 1)
             bVinculoBusca1 = True
             Dim sVinSql As String
             sVinSql = "UPDATE VinculoXMLProduto SET " & _
                       "CustoUnitario = " & FSQL(xVUnCom / qtfracao) & ", " & _
                       "DataAtualizacao = GETDATE(), " & _
                       "NCM = '" & sVinNCM & "', " & _
                       "CFOP = " & sVinCFOP & ", " & _
                       "CST = '" & Replace(csticms, "'" , "''") & "', " & _
                       "pICMS = " & FSQL(ICMSAliquota) & ", " & _
                       "IPICST = '" & cstipi & "', " & _
                       "IPIpIPI = " & FSQL(IPIpIPI) & ", " & _
                       "PISCST = '" & cstpis & "', "
             sVinSql = sVinSql & _
                       "PISpPIS = " & FSQL(PISpPIS) & ", " & _
                       "COFINSCST = '" & cstcofins & "', " & _
                       "COFINSpCOFINS = " & FSQL(COFINSpCOFINS) & ", " & _
                       "CEST = '" & sVinCEST & "' " & _
                       "WHERE IDFornecedor = " & Codigo_fornecedor & " AND cProd = '" & Replace(xCProd, "'" , "''") & "'"
             SQLExecuta sVinSql
          End If
          TbVinculo.Close

          '--- Busca 2: VinculoXMLProduto por EANEmbalagem ---
          If CodProduto = 0 And Not Vazio(cEAN) Then
             RsOpen TbVinculo, "SELECT IDProduto, Fracionamento FROM VinculoXMLProduto " & _
                               " WHERE EANEmbalagem = '" & cEAN & "'"
             If TbVinculo.RecordCount > 0 Then
                CodProduto = TbVinculo!IDProduto
                qtfracao = IIf(TbVinculo!Fracionamento > 0, TbVinculo!Fracionamento, 1)
             End If
             TbVinculo.Close
          End If

          '--- Busca 3: Produtos.EAN por cEAN (EAN da embalagem comercial) ---
          If CodProduto = 0 And Not Vazio(cEAN) Then
             CodProduto = SQLExecutaRetorno("SELECT Codigo FROM Produtos WHERE EAN = '" & cEAN & "'", "Codigo", 0)
          End If

          '--- Busca 4: Produtos.EAN por cEANTrib (EAN da unidade = EAN interno, ex: fornecedor envia caixa mas informa EAN da unidade em cEANTrib) ---
          If CodProduto = 0 And Not Vazio(cEANTrib) Then
             CodProduto = SQLExecutaRetorno("SELECT Codigo FROM Produtos WHERE EAN = '" & cEANTrib & "'", "Codigo", 0)
             If CodProduto > 0 And cEANTrib <> cEAN And xQCom > 0 And xQTrib > xQCom Then
                'cEANTrib diferente de cEAN: fornecedor esta enviando embalagem (caixa)
                'fracionamento = quantas unidades tributaveis ha por unidade comercial
                qtfracao = xQTrib / xQCom
             End If
          End If

          '--- Auto-vincular: INSERT em VinculoXMLProduto quando produto foi encontrado mas nao havia vinculo ---
          If CodProduto > 0 And Not bVinculoBusca1 Then
             Dim sVinxProd As String, sVinuCom As String, sVinEANProduto As String
             Dim sAutoVin As String
             Dim oDetAuto As Object
             sVinxProd = ""
             sVinuCom = ""
             sVinEANProduto = ""
             Set oDetAuto = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i)
             If oDetAuto.selectNodes("prod/xProd").Length > 0 Then sVinxProd = oDetAuto.selectNodes("prod/xProd").item(0).Text
             If oDetAuto.selectNodes("prod/uCom").Length > 0 Then sVinuCom = oDetAuto.selectNodes("prod/uCom").item(0).Text
             Set oDetAuto = Nothing
             sVinEANProduto = SQLExecutaRetorno("SELECT ISNULL(EAN,'') r FROM Produtos WHERE Codigo = " & CodProduto, "r", "")
             sAutoVin = "IF NOT EXISTS (SELECT 1 FROM VinculoXMLProduto WHERE IDFornecedor = " & Codigo_fornecedor & _
                        " AND cProd = '" & Replace(xCProd, "'" , "''") & "') " & _
                        "INSERT INTO VinculoXMLProduto (IDFornecedor,cProd,EANEmbalagem,xProd,uCom," & _
                        "NCM,CFOP,CST,pICMS,IPICST,IPIpIPI,PISCST,PISpPIS,COFINSCST,COFINSpCOFINS,CEST," & _
                        "IDProduto,EANProduto,Fracionamento,CustoUnitario,DataAtualizacao) VALUES (" & _
                        Codigo_fornecedor & ",'" & Replace(xCProd, "'" , "''") & "','" & cEAN & "',"
             sAutoVin = sAutoVin & _
                        "'" & Replace(sVinxProd, "'" , "''") & "','" & sVinuCom & "'," & _
                        "'" & sVinNCM & "'," & sVinCFOP & ",'" & Replace(csticms, "'" , "''") & "'," & _
                        FSQL(ICMSAliquota) & ",'" & cstipi & "'," & FSQL(IPIpIPI) & ",'" & cstpis & "',"
             sAutoVin = sAutoVin & _
                        FSQL(PISpPIS) & ",'" & cstcofins & "'," & FSQL(COFINSpCOFINS) & ",'" & sVinCEST & "'," & _
                        CodProduto & ",'" & sVinEANProduto & "'," & FSQL(qtfracao) & "," & _
                        FSQL(xVUnCom / qtfracao) & ",GETDATE())"
             SQLExecuta sAutoVin
          End If
    
           xCfop = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/CFOP").item(0).Text
          ''If CodProduto > 0 Then
          ''   'PTab Produtos, "IDProduto", CodProduto
          ''    xCfop = VerificaCFOPEntrada(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/dest/enderDest/UF").Item(0).Text), UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").Item(0).Text), csticms, Produtos!CfopEnt)
          '' Else
          ''    xCfop = VerificaCFOPEntrada(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/dest/enderDest/UF").Item(0).Text), UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").Item(0).Text), csticms, "")
          ''End If
           'Compras_XML!CFOP = xCfop
           Compras_XML.Update
           
           'Lançando os itens no Banco
           Compras_Itens_XML.AddNew
           Compras_Itens_XML!CodigoNota = Compra
           'Compras_Itens_XML!Ite_documento = Compra
           Compras_Itens_XML!item = SQLExecutaRetorno("SELECT ISNULL(MAX(Item),0) r FROM EntradaEstoqueItens WHERE CodigoNota = " & Compra, "r", 0) + 1
           Compras_Itens_XML!CodigoProduto = CodProduto
           Compras_Itens_XML!Referencia = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/cProd").item(0).Text
           Compras_Itens_XML!Genero = Codigo_fornecedor
           'Compras_itens_xml!Barras = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cEAN").Item(0).Text
           Compras_Itens_XML!NomeProduto = UCase(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/xProd").item(0).Text)          'xProd:    Descricao do produto ou servico <xProd>
           Compras_Itens_XML!NCM = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/NCM").item(0).Text                              'NCM:      Codigo NCM (Nomenclatura Comum do Mercosul) <NCM>
           Compras_Itens_XML!CEST = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/CEST").item(0).Text                            'CEST:     Codigo Especificador da Substituicao Tributaria <CEST>
           Compras_Itens_XML!indEscala = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/indEscala").item(0).Text                  'indEscala:Indicador de Producao em Escala Relevante: S=Sim, N=Nao <indEscala>
           'Compras_itens_xml!Ex_ipi = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/EXTIPI").Item(0).Text
           Compras_Itens_XML!CFOP = xCfop                                                                                                                  'CFOP:     Codigo Fiscal de Operacoes e Prestacoes <CFOP>
           Compras_Itens_XML!UnidadeComercial = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/uCom").item(0).Text                 'uCom:     Unidade comercial do produto (ex: UN, CX, KG) <uCom>
           Compras_Itens_XML!QuantidadeComercial = (CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/qCom").item(0).Text, ".", ",", UM_A_UM)) * qtfracao)          'qCom: Quantidade comercial <qCom>
           Compras_Itens_XML!ValorUnitarioComercializacao = (CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/vUnCom").item(0).Text, ".", ",", UM_A_UM)) / qtfracao) 'vUnCom: Valor unitario de comercializacao <vUnCom>
           Compras_Itens_XML!ValorTotalBruto = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/vProd").item(0).Text, ".", ",", UM_A_UM))                           'vProd: Valor total bruto dos produtos/servicos <vProd>
           Compras_Itens_XML!indTot = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/indTot").item(0).Text                        'indTot:   Indica se o item compoe o valor total da NF-e: 0=Nao compoe, 1=Compoe <indTot>
           If XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/nFCI").Length > 0 Then                                               'nFCI:     Numero da FCI - Ficha de Conteudo de Importacao (opcional, apenas para importados com ST) <nFCI>
              Compras_Itens_XML!nFCI = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/nFCI").item(0).Text
           Else
              Compras_Itens_XML!nFCI = ""
           End If
           'Compras_Itens_XML!Ite_frete = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vFrete").Item(0).Text, ".", ",", UM_A_UM))
           'Compras_itens_xml!Vlr_seguro = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vSeg").Item(0).Text
           ''Compras_Itens_XML!ValorDesconto = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vDesc").Item(0).Text, ".", ",", UM_A_UM))
           If XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/vOutro").Length > 0 Then  'vOutro: Valor de outras despesas acessorias do item <vOutro> (opcional)
              Compras_Itens_XML!vOutro = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/vOutro").item(0).Text, ".", ",", UM_A_UM))
           Else
              Compras_Itens_XML!vOutro = 0
           End If
           'Compras_itens_xml!Origem = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").Item(i).childNodes(1).childNodes(0).childNodes(0).selectNodes("orig").Item(0).Text
           '==============================================================
           ' GRAVACAO - ICMS
           '==============================================================
           Compras_Itens_XML!CST = csticms          'orig+CST (ou CSOSN): origem e situacao tributaria <orig><CST>/<CSOSN>
           Compras_Itens_XML!modBC = ICMSmodBC      'modBC:   Modalidade de determinacao da base de calculo <modBC>
           Compras_Itens_XML!vBC = ICMSBase          'vBC:     Base de calculo do ICMS <vBC>
           Compras_Itens_XML!pICMS = ICMSAliquota    'pICMS:   Aliquota do ICMS (%) <pICMS>
           Compras_Itens_XML!vICMS = ICMSValor       'vICMS:   Valor do ICMS <vICMS>
           Compras_Itens_XML!pRedBC = ICMSpRedBC         'pRedBC:     Percentual de reducao da base de calculo <pRedBC> (ICMS20/51/70/90)
           Compras_Itens_XML!vICMSOp = ICMSvICMSOp       'vICMSOp:    Valor do ICMS da operacao antes do diferimento <vICMSOp> (ICMS51)
           Compras_Itens_XML!pDif = ICMSpDif              'pDif:       Percentual do diferimento <pDif> (ICMS51)
           Compras_Itens_XML!vICMSDeson = ICMSvICMSDeson  'vICMSDeson: Valor do ICMS desonerado <vICMSDeson> (ICMS40/41)
           Compras_Itens_XML!motDesICMS = ICMSmotDesICMS  'motDesICMS: Motivo da desoneracao do ICMS <motDesICMS> (ICMS40/41)
           '--- ICMS-ST ---
           Compras_Itens_XML!modBCST = ICMSmodBCST   'modBCST: Modalidade da base de calculo do ICMS ST <modBCST> (ICMS10/70)
           Compras_Itens_XML!pRedBCST = ICMSpRedBCST 'pRedBCST:Percentual de reducao da base de calculo ST <pRedBCST> (ICMS10)
           Compras_Itens_XML!pMVAST = ICMSpMVAST     'pMVAST:  Percentual da Margem de Valor Adicionado ST <pMVAST> (ICMS70/90)
           Compras_Itens_XML!vBCST = ICMSvBCST       'vBCST:   Base de calculo do ICMS ST <vBCST> (ICMS10/70/90)
           Compras_Itens_XML!pICMSST = ICMSpICMSST   'pICMSST: Aliquota do ICMS ST (%) <pICMSST> (ICMS10/70)
           Compras_Itens_XML!vICMSST = ICMSvICMSST       'vICMSST:    Valor do ICMS ST <vICMSST> (ICMS10/70/90)
           Compras_Itens_XML!vBCFCPST = ICMSvBCFCPST     'vBCFCPST:   Base de calculo do FCP retido por ST <vBCFCPST> (ICMS10/70/90)
           Compras_Itens_XML!pFCPST = ICMSpFCPST          'pFCPST:     Percentual do FCP retido por ST <pFCPST> (ICMS10/70/90)
           Compras_Itens_XML!vFCPST = ICMSvFCPST          'vFCPST:     Valor do FCP retido por ST <vFCPST> (ICMS10/70/90)
           '--- ICMS60 - ST Retido anteriormente ---
           Compras_Itens_XML!vBCSTRet = ICMSvBCSTRet               'vBCSTRet:       Base de calculo do ICMS ST retido anteriormente <vBCSTRet> (ICMS60)
           Compras_Itens_XML!pST = ICMSpST                          'pST:            Aliquota suportada pelo consumidor final <pST> (ICMS60)
           Compras_Itens_XML!vICMSSubstituto = ICMSvICMSSubstituto  'vICMSSubstituto:Valor do ICMS do substituto tributario <vICMSSubstituto> (ICMS60)
           Compras_Itens_XML!vICMSSTRet = ICMSvICMSSTRet            'vICMSSTRet:     Valor do ICMS ST retido anteriormente <vICMSSTRet> (ICMS60)
           Compras_Itens_XML!vBCFCPSTRet = ICMSvBCFCPSTRet          'vBCFCPSTRet:    Base de calculo do FCP ST retido <vBCFCPSTRet> (ICMS60)
           Compras_Itens_XML!pFCPSTRet = ICMSpFCPSTRet              'pFCPSTRet:      Percentual do FCP ST retido <pFCPSTRet> (ICMS60)
           Compras_Itens_XML!vFCPSTRet = ICMSvFCPSTRet              'vFCPSTRet:      Valor do FCP ST retido <vFCPSTRet> (ICMS60)

           '==============================================================
           ' GRAVACAO - IPI
           '==============================================================
           Compras_Itens_XML!IPIcEnq = IPIcEnq      'cEnq:     Codigo de enquadramento legal do IPI <cEnq>
           Compras_Itens_XML!IPIclEnq = ""           'clEnq:    Classe de enquadramento do IPI (ex.: 001) <clEnq> - nao presente na maioria das NF-e
           Compras_Itens_XML!IPICNPJProd = ""        'CNPJProd: CNPJ do produtor da mercadoria <CNPJProd> - apenas para industrializacao por encomenda
           Compras_Itens_XML!IPIcSelo = ""           'cSelo:    Codigo do Selo de Controle do IPI <cSelo> - apenas para cigarros/bebidas
           Compras_Itens_XML!IPIqSelo = 0            'qSelo:    Quantidade de Selos de Controle <qSelo>
           Compras_Itens_XML!IPICST = cstipi         'CST:      Codigo de Situacao Tributaria do IPI <CST>
           Compras_Itens_XML!IPIvBC = IPIvBC         'vBC:      Base de calculo do IPI <vBC>
           Compras_Itens_XML!IPIpIPI = IPIpIPI       'pIPI:     Aliquota do IPI (%) <pIPI>
           Compras_Itens_XML!IPIvIPI = IPIvIPI       'vIPI:     Valor do IPI <vIPI>
           Compras_Itens_XML!IPIvUnid = 0            'vUnid:    Valor por Unidade Tributavel (IPI por qtde) <vUnid>
           Compras_Itens_XML!IPIqUnid = 0            'qUnid:    Quantidade total na unidade padrao para tributacao <qUnid>

           '==============================================================
           ' GRAVACAO - PIS
           '==============================================================
           Compras_Itens_XML!pisCST = pisCST              'CST:       Codigo de Situacao Tributaria do PIS <CST>
           Compras_Itens_XML!PISvBC = PISvBC              'vBC:       Base de calculo do PIS <vBC> (CST 01/02/99)
           Compras_Itens_XML!PISpPIS = PISpPIS            'pPIS:      Aliquota do PIS em percentual <pPIS> (CST 01/02/99)
           Compras_Itens_XML!PISvPIS = PISvPIS            'vPIS:      Valor do PIS <vPIS>
           Compras_Itens_XML!PISqBCProd = PISqBCProd      'qBCProd:   Quantidade vendida - base de calculo por qtde <qBCProd> (CST 03)
           Compras_Itens_XML!PISvAliqProd = PISvAliqProd  'vAliqProd: Aliquota do PIS em reais por unidade de produto <vAliqProd> (CST 03)

           '==============================================================
           ' GRAVACAO - COFINS
           '==============================================================
           Compras_Itens_XML!cofinsCST = cofinsCST             'CST:       Codigo de Situacao Tributaria do COFINS <CST>
           Compras_Itens_XML!COFINSvBC = COFINSvBC             'vBC:       Base de calculo do COFINS <vBC> (CST 01/02/99)
           Compras_Itens_XML!COFINSpCOFINS = COFINSpCOFINS     'pCOFINS:   Aliquota do COFINS em percentual <pCOFINS> (CST 01/02/99)
           Compras_Itens_XML!COFINSvCOFINS = COFINSvCOFINS     'vCOFINS:   Valor do COFINS <vCOFINS>
           Compras_Itens_XML!COFINSqBCProd = COFINSqBCProd     'qBCProd:   Quantidade vendida - base de calculo por qtde <qBCProd> (CST 03)
           Compras_Itens_XML!COFINSvAliqProd = COFINSvAliqProd 'vAliqProd: Aliquota do COFINS em reais por unidade de produto <vAliqProd> (CST 03)

           '==============================================================
           ' GRAVACAO - Dados do Produto (campos adicionais)
           '==============================================================
           Compras_Itens_XML!EANTrib = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/cEANTrib").item(0).Text                                                                           'cEANTrib:  GTIN (EAN) da unidade tributavel <cEANTrib>
           Compras_Itens_XML!UnidadeTributavel = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/uTrib").item(0).Text                                                                    'uTrib:     Unidade tributavel (ex: CX, UN, KG) <uTrib>
           Compras_Itens_XML!QuantidadeTributavel = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/qTrib").item(0).Text, ".", ",", UM_A_UM))                             'qTrib:     Quantidade na unidade tributavel <qTrib>
           Compras_Itens_XML!ValorUnitarioTributario = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/vUnTrib").item(0).Text, ".", ",", UM_A_UM))                        'vUnTrib:   Valor unitario de tributacao <vUnTrib>
           Compras_Itens_XML!ValorFrete = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/vFrete").item(0).Text, ".", ",", UM_A_UM))                                      'vFrete:    Valor do frete rateado no item <vFrete>
           Compras_Itens_XML!ValorSeguro = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/vSeg").item(0).Text, ".", ",", UM_A_UM))                                       'vSeg:      Valor do seguro rateado no item <vSeg>
           Compras_Itens_XML!ValorDesconto = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/vDesc").item(0).Text, ".", ",", UM_A_UM))                                    'vDesc:     Valor do desconto concedido no item <vDesc>
           Compras_Itens_XML!Desconto = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("prod/vDesc").item(0).Text, ".", ",", UM_A_UM))                                         'vDesc:     Mesmo valor de ValorDesconto - na NF-e o desconto e sempre em valor fixo
           Compras_Itens_XML!TipoDesconto = 1                                                                                                                                                                      'TipoDesconto: 1 = desconto em valor fixo (padrao NF-e, tag <vDesc> sempre em reais)
           Compras_Itens_XML!InformacoesAdicionaisProduto = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("infAdProd").item(0).Text                                                          'infAdProd: Informacoes adicionais do produto (texto livre) <infAdProd>
           Compras_Itens_XML!ValorTributos = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/vTotTrib").item(0).Text, ".", ",", UM_A_UM))                              'vTotTrib:  Valor aproximado total dos tributos (IBPT) <vTotTrib>
           Compras_Itens_XML!ItemXML = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).Attributes.getNamedItem("nItem").Text                                                                               'nItem:     Numero do item no XML da NF-e (atributo do elemento <det>)

           '==============================================================
           ' GRAVACAO - Imposto de Importacao (II)
           '==============================================================
           Compras_Itens_XML!IIvBC = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/II/vBC").item(0).Text, ".", ",", UM_A_UM))           'vBC:     Base de calculo do Imposto de Importacao <vBC>
           Compras_Itens_XML!IIvDespAdu = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/II/vDespAdu").item(0).Text, ".", ",", UM_A_UM)) 'vDespAdu:Valor das despesas aduaneiras <vDespAdu>
           Compras_Itens_XML!IIvII = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/II/vII").item(0).Text, ".", ",", UM_A_UM))           'vII:     Valor do Imposto de Importacao <vII>
           Compras_Itens_XML!IIvIOF = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/II/vIOF").item(0).Text, ".", ",", UM_A_UM))         'vIOF:    Valor do IOF (apenas para importacao financiada) <vIOF>

           '==============================================================
           ' GRAVACAO - ISSQN (Imposto Sobre Servicos - notas de servico)
           '==============================================================
           Compras_Itens_XML!ISSvBC = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/ISSQN/vBC").item(0).Text, ".", ",", UM_A_UM))             'vBC:       Base de calculo do ISSQN <vBC>
           Compras_Itens_XML!ISSvAliq = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/ISSQN/vAliq").item(0).Text, ".", ",", UM_A_UM))         'vAliq:     Aliquota do ISSQN <vAliq>
           Compras_Itens_XML!ISSvISSQN = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/ISSQN/vISSQN").item(0).Text, ".", ",", UM_A_UM))       'vISSQN:    Valor do ISSQN <vISSQN>
           Compras_Itens_XML!ISScMunFG = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/ISSQN/cMunFG").item(0).Text                                            'cMunFG:    Codigo do municipio de ocorrencia do fato gerador do ISSQN <cMunFG>
           Compras_Itens_XML!ISScListServ = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/ISSQN/cListServ").item(0).Text                                      'cListServ: Codigo do servico conforme lista do ISSQN <cListServ>
           Compras_Itens_XML!EAN = cEAN

           '==============================================================
           ' GRAVACAO - IBSCBS (Reforma Tributaria - IBS+CBS por item)
           '==============================================================
           Compras_Itens_XML!IBSCBSCST = ""
           Compras_Itens_XML!IBSCBScClassTrib = ""
           Compras_Itens_XML!IBSCBSvBC = 0
           Compras_Itens_XML!IBSUFpAliq = 0
           Compras_Itens_XML!IBSUFpRedAliq = 0
           Compras_Itens_XML!IBSUFpAliqEfet = 0
           Compras_Itens_XML!IBSUFvIBS = 0
           Compras_Itens_XML!IBSMunpAliq = 0
           Compras_Itens_XML!IBSMunpRedAliq = 0
           Compras_Itens_XML!IBSMunpAliqEfet = 0
           Compras_Itens_XML!IBSMunvIBS = 0
           Compras_Itens_XML!IBSvIBS = 0
           Compras_Itens_XML!CBSpAliq = 0
           Compras_Itens_XML!CBSpRedAliq = 0
           Compras_Itens_XML!CBSpAliqEfet = 0
           Compras_Itens_XML!CBSvCBS = 0
           Compras_Itens_XML!IBSUFpDif = 0
           Compras_Itens_XML!IBSUFvDif = 0
           Compras_Itens_XML!IBSUFvDevTrib = 0
           Compras_Itens_XML!IBSMunpDif = 0
           Compras_Itens_XML!IBSMunvDif = 0
           Compras_Itens_XML!IBSMunvDevTrib = 0
           Compras_Itens_XML!CBSpDif = 0
           Compras_Itens_XML!CBSvDif = 0
           Compras_Itens_XML!CBSvDevTrib = 0
           Compras_Itens_XML!IBSCBSMonoqBCMono = 0
           Compras_Itens_XML!IBSCBSMonoadRemIBS = 0
           Compras_Itens_XML!IBSCBSMonoadRemCBS = 0
           Compras_Itens_XML!IBSCBSMonovIBSMono = 0
           Compras_Itens_XML!IBSCBSMonovCBSMono = 0
           Compras_Itens_XML!IBSCBSMonoqBCMonoReten = 0
           Compras_Itens_XML!IBSCBSMonoadRemIBSReten = 0
           Compras_Itens_XML!IBSCBSMonovIBSMonoReten = 0
           Compras_Itens_XML!IBSCBSMonoadRemCBSReten = 0
           Compras_Itens_XML!IBSCBSMonovCBSMonoReten = 0
           Compras_Itens_XML!IBSCBSMonoqBCMonoRet = 0
           Compras_Itens_XML!IBSCBSMonoadRemIBSRet = 0
           Compras_Itens_XML!IBSCBSMonovIBSMonoRet = 0
           Compras_Itens_XML!IBSCBSMonoadRemCBSRet = 0
           Compras_Itens_XML!IBSCBSMonovCBSMonoRet = 0
           Compras_Itens_XML!IBSCBSMonovTotIBSMonoItem = 0
           Compras_Itens_XML!IBSCBSMonovTotCBSMonoItem = 0
           Dim oIBSCBS As Object
           Set oIBSCBS = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/IBSCBS").item(0)
           If Not oIBSCBS Is Nothing Then
              Compras_Itens_XML!IBSCBSCST = oIBSCBS.selectNodes("CST").item(0).Text
              If oIBSCBS.selectNodes("cClassTrib").Length > 0 Then
                 Compras_Itens_XML!IBSCBScClassTrib = oIBSCBS.selectNodes("cClassTrib").item(0).Text
              End If
              If oIBSCBS.selectNodes("gIBSCBS").Length > 0 Then
                 Compras_Itens_XML!IBSCBSvBC = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/vBC").item(0).Text, ".", ",", UM_A_UM))
                 Compras_Itens_XML!IBSUFpAliq = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSUF/pIBSUF").item(0).Text, ".", ",", UM_A_UM))
                 Compras_Itens_XML!IBSUFvIBS = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSUF/vIBSUF").item(0).Text, ".", ",", UM_A_UM))
                 If oIBSCBS.selectNodes("gIBSCBS/gIBSUF/gDif").Length > 0 Then
                    Compras_Itens_XML!IBSUFpDif = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSUF/gDif/pDif").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSUFvDif = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSUF/gDif/vDif").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oIBSCBS.selectNodes("gIBSCBS/gIBSUF/gDevTrib").Length > 0 Then
                    Compras_Itens_XML!IBSUFvDevTrib = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSUF/gDevTrib/vDevTrib").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oIBSCBS.selectNodes("gIBSCBS/gIBSUF/gRed").Length > 0 Then
                    Compras_Itens_XML!IBSUFpRedAliq = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSUF/gRed/pRedAliq").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSUFpAliqEfet = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSUF/gRed/pAliqEfet").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 Compras_Itens_XML!IBSMunpAliq = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSMun/pIBSMun").item(0).Text, ".", ",", UM_A_UM))
                 Compras_Itens_XML!IBSMunvIBS = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSMun/vIBSMun").item(0).Text, ".", ",", UM_A_UM))
                 If oIBSCBS.selectNodes("gIBSCBS/gIBSMun/gDif").Length > 0 Then
                    Compras_Itens_XML!IBSMunpDif = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSMun/gDif/pDif").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSMunvDif = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSMun/gDif/vDif").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oIBSCBS.selectNodes("gIBSCBS/gIBSMun/gDevTrib").Length > 0 Then
                    Compras_Itens_XML!IBSMunvDevTrib = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSMun/gDevTrib/vDevTrib").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oIBSCBS.selectNodes("gIBSCBS/gIBSMun/gRed").Length > 0 Then
                    Compras_Itens_XML!IBSMunpRedAliq = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSMun/gRed/pRedAliq").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSMunpAliqEfet = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gIBSMun/gRed/pAliqEfet").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 Compras_Itens_XML!IBSvIBS = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/vIBS").item(0).Text, ".", ",", UM_A_UM))
                 Compras_Itens_XML!CBSpAliq = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gCBS/pCBS").item(0).Text, ".", ",", UM_A_UM))
                 Compras_Itens_XML!CBSvCBS = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gCBS/vCBS").item(0).Text, ".", ",", UM_A_UM))
                 If oIBSCBS.selectNodes("gIBSCBS/gCBS/gDif").Length > 0 Then
                    Compras_Itens_XML!CBSpDif = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gCBS/gDif/pDif").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!CBSvDif = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gCBS/gDif/vDif").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oIBSCBS.selectNodes("gIBSCBS/gCBS/gDevTrib").Length > 0 Then
                    Compras_Itens_XML!CBSvDevTrib = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gCBS/gDevTrib/vDevTrib").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oIBSCBS.selectNodes("gIBSCBS/gCBS/gRed").Length > 0 Then
                    Compras_Itens_XML!CBSpRedAliq = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gCBS/gRed/pRedAliq").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!CBSpAliqEfet = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBS/gCBS/gRed/pAliqEfet").item(0).Text, ".", ",", UM_A_UM))
                 End If
              End If
              ' IBSCBS Monofasico (gIBSCBSMono) - combustiveis e monofasicos (NT 2025.002 UB84-105)
              If oIBSCBS.selectNodes("gIBSCBSMono").Length > 0 Then
                 If oIBSCBS.selectNodes("gIBSCBSMono/gMonoPadrao").Length > 0 Then
                    Compras_Itens_XML!IBSCBSMonoqBCMono = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoPadrao/qBCMono").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonoadRemIBS = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoPadrao/adRemIBS").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonoadRemCBS = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoPadrao/adRemCBS").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonovIBSMono = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoPadrao/vIBSMono").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonovCBSMono = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoPadrao/vCBSMono").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oIBSCBS.selectNodes("gIBSCBSMono/gMonoReten").Length > 0 Then
                    Compras_Itens_XML!IBSCBSMonoqBCMonoReten = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoReten/qBCMonoReten").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonoadRemIBSReten = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoReten/adRemIBSReten").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonovIBSMonoReten = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoReten/vIBSMonoReten").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonoadRemCBSReten = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoReten/adRemCBSReten").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonovCBSMonoReten = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoReten/vCBSMonoReten").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 If oIBSCBS.selectNodes("gIBSCBSMono/gMonoRet").Length > 0 Then
                    Compras_Itens_XML!IBSCBSMonoqBCMonoRet = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoRet/qBCMonoRet").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonoadRemIBSRet = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoRet/adRemIBSRet").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonovIBSMonoRet = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoRet/vIBSMonoRet").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonoadRemCBSRet = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoRet/adRemCBSRet").item(0).Text, ".", ",", UM_A_UM))
                    Compras_Itens_XML!IBSCBSMonovCBSMonoRet = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/gMonoRet/vCBSMonoRet").item(0).Text, ".", ",", UM_A_UM))
                 End If
                 Compras_Itens_XML!IBSCBSMonovTotIBSMonoItem = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/vTotIBSMonoItem").item(0).Text, ".", ",", UM_A_UM))
                 Compras_Itens_XML!IBSCBSMonovTotCBSMonoItem = CDbl(Substitui(oIBSCBS.selectNodes("gIBSCBSMono/vTotCBSMonoItem").item(0).Text, ".", ",", UM_A_UM))
              End If
           End If
           Set oIBSCBS = Nothing

           '==============================================================
           ' GRAVACAO - IS (Imposto Seletivo - NT 2025.002 - Grupo UB01-UB11)
           '==============================================================
           Compras_Itens_XML!ISCST = ""
           Compras_Itens_XML!IScClassTrib = ""
           Compras_Itens_XML!ISvBCIS = 0
           Compras_Itens_XML!ISpIS = 0
           Compras_Itens_XML!ISvIS = 0
           Dim oIS As Object
           Set oIS = XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("imposto/IS").item(0)
           If Not oIS Is Nothing Then
              Compras_Itens_XML!ISCST = oIS.selectNodes("CSTIS").item(0).Text
              If oIS.selectNodes("cClassTribIS").Length > 0 Then
                 Compras_Itens_XML!IScClassTrib = oIS.selectNodes("cClassTribIS").item(0).Text
              End If
              If oIS.selectNodes("vBCIS").Length > 0 Then
                 Compras_Itens_XML!ISvBCIS = CDbl(Substitui(oIS.selectNodes("vBCIS").item(0).Text, ".", ",", UM_A_UM))
              End If
              If oIS.selectNodes("pIS").Length > 0 Then
                 Compras_Itens_XML!ISpIS = CDbl(Substitui(oIS.selectNodes("pIS").item(0).Text, ".", ",", UM_A_UM))
              End If
              If oIS.selectNodes("vIS").Length > 0 Then
                 Compras_Itens_XML!ISvIS = CDbl(Substitui(oIS.selectNodes("vIS").item(0).Text, ".", ",", UM_A_UM))
              End If
           End If
           Set oIS = Nothing

           '==============================================================
           ' GRAVACAO - vItem (valor total liquido do item)
           '==============================================================
           Compras_Itens_XML!vItem = 0
           If XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("vItem").Length > 0 Then
              Compras_Itens_XML!vItem = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/det").item(i).selectNodes("vItem").item(0).Text, ".", ",", UM_A_UM))
           End If

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
    '------------- FIM DA IMPORTAaaO DOS ITENS ------------------

    '------------- VINCULACAO DE PRODUTOS NAO IDENTIFICADOS ---------
    'Verifica se restaram itens com CodigoProduto = 0 e abre o formulario de vinculacao
    Dim qtdNaoIdentificados As Long
    qtdNaoIdentificados = SQLExecutaRetorno( _
        "SELECT COUNT(*) r FROM EntradaEstoqueItens " & _
        "WHERE CodigoNota = " & Compra & " AND (CodigoProduto = 0 OR CodigoProduto IS NULL)", "r", 0)

    If qtdNaoIdentificados > 0 Then
       If MsgBox(qtdNaoIdentificados & " produto(s) desta nota nao foram identificados no cadastro." & vbCrLf & vbCrLf & _
                 "Deseja vincular agora?", vbQuestion + vbYesNo, "Produtos nao identificados") = vbYes Then
          Dim frmVinc2 As New frmVinculoProdutoXML
          frmVinc2.NumeroEntrada = Compra
          frmVinc2.Show vbModal
          Unload frmVinc2
          Set frmVinc2 = Nothing
       End If
    End If
    '------------- FIM DA VINCULACAO ---------

    '------------- LANCANDO DUPLICATAS NO A_PAGAR ---------

    ' Forma de pagamento real a partir de <pag/detPag/tPag> (NF-e 4.0+)
    Dim sFormaPag As String
    sFormaPag = "OUTROS"
    If XMLDOC.selectNodes("nfeProc/NFe/infNFe/pag/detPag/tPag").Length > 0 Then
       Select Case Trim(XMLDOC.selectNodes("nfeProc/NFe/infNFe/pag/detPag/tPag").item(0).Text)
          Case "01": sFormaPag = "DINHEIRO"
          Case "02": sFormaPag = "CHEQUE"
          Case "03": sFormaPag = "CARTAO CREDITO"
          Case "04": sFormaPag = "CARTAO DEBITO"
          Case "05": sFormaPag = "CREDITO LOJA"
          Case "10": sFormaPag = "VALE ALIMENTACAO"
          Case "11": sFormaPag = "VALE REFEICAO"
          Case "12": sFormaPag = "VALE PRESENTE"
          Case "13": sFormaPag = "VALE COMBUSTIVEL"
          Case "14": sFormaPag = "DUPLICATA MERCANTIL"
          Case "15": sFormaPag = "BOLETO"
          Case "16": sFormaPag = "DEPOSITO BANCARIO"
          Case "17": sFormaPag = "PIX"
          Case "18": sFormaPag = "TRANSF. BANCARIA"
          Case "19": sFormaPag = "FIDELIDADE"
          Case "90": sFormaPag = "SEM PAGAMENTO"
          Case Else:  sFormaPag = "OUTROS"
       End Select
    End If

    ' Lanca duplicatas no a_pagar somente se for pagamento a prazo:
    '   sIndPag = "1"  => prazo (NF-e 4.0+ ou 3.x com indPag explicito)
    '   sIndPag = ""   => tag ausente na XML (NF-e antigas) - lanca por retrocompatibilidade
    Dim bPagPrazo As Boolean
    bPagPrazo = (sIndPag = "1") Or (sIndPag = "")

    i = 0
    qtdDup = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/cobr/dup").Length 'Contando quantas parcelas existem

    If qtdDup > 0 And bPagPrazo Then
       'Nome do fornecedor: limitado a 40 chars (tamanho da coluna FAVORECIDO em a_pagar)
       sFavorecido = Replace(Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xNome").item(0).Text), 40), "'", "''")
       For i = 0 To qtdDup - 1
          'nDup: Numero da duplicata - tag <nDup>
          sDupNum = XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").item(i).selectNodes("nDup").item(0).Text
          'dVenc: Data de vencimento da duplicata - tag <dVenc>
          sDupVenc = XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").item(i).selectNodes("dVenc").item(0).Text
          'vDup: Valor da duplicata - tag <vDup>
          vDupValor = CDbl(Substitui(XMLDOC.selectNodes("nfeProc/NFe/infNFe/cobr/dup").item(i).selectNodes("vDup").item(0).Text, ".", ",", UM_A_UM))

          str_sql = "INSERT INTO a_pagar " & _
                    "(CODIGO, CADASTRO, VENCIMENTO, STATUS, DESCRICAO, FAVORECIDO, SETOR, TIPO, DATA_PGTO, FORMA, FONTE, VALOR_UND, HAVERES, QUANT, TOTAL, REF) " & _
                    "VALUES (" & _
                    "(SELECT ISNULL(MAX(CODIGO),0)+1 FROM a_pagar), " & _
                    "'" & Format(Date, "yyyymmdd") & "', " & _
                    "'" & Replace(Trim(sDupVenc), "-", "") & "', " & _
                    "'A PAGAR', " & _
                    "'NF-e " & LPad(Compra, 6, "0") & " - DUP: " & sDupNum & "', " & _
                    "'" & sFavorecido & "', " & _
                    "'EMPRESA', " & _
                    "'TEMPORARIA', " & _
                    "NULL, " & _
                    "'" & sFormaPag & "', " & _
                    "'ENTRADA XML', " & _
                    FSQL(vDupValor) & ", " & _
                    "0, NULL, " & _
                    FSQL(vDupValor) & ", " & _
                    Compra & ")"

          msgResultado = SQLExecuta(str_sql)
          If Not Vazio(msgResultado) Then
             MsgBox "Erro ao lancar duplicata " & sDupNum & " no a_pagar: " & msgResultado, vbCritical
             ImportarXML = False: Exit Function
          End If
       Next i
    End If
    '------------- FIM DO LANCAMENTO DAS DUPLICATAS ------------------

pulaParcelas:
   MsgBox "NF-e importada com sucesso!", vbInformation + vbOKOnly, "Importa XML"

   str_sql = "SELECT Count(*) r FROM EntradaEstoqueItens WHERE CodigoNota = " & Compra & " AND CodigoProduto = 0"
   If SQLExecutaRetorno(str_sql, "r", 0) > 0 Then
       xCaminhoXML = arqXML
   End If
   
   RsOpen TbEntrada, "SELECT *, " & _
                     "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status " & _
                     "FROM EntradaEstoque WHERE CodigoNota = " & Compra
   Load_Controls
   Frm_NF.Tab = 0

   ImportarXML = True
End Function

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
varRAZAOForn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xNome").item(0).Text), 35)
varFantasiaFor = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xFant").item(0).Text), 35)
If tipoEmp = 0 Then varCNPJforn = CPF
If tipoEmp = 1 Then varCNPJforn = CNPJ
varIEforn = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/IE").item(0).Text
varENDERECOforn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xLgr").item(0).Text), 40) & ", " & Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").item(0).Text), 8)
varBairroforn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xBairro").item(0).Text), 20)
varCidadeforn = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xMun").item(0).Text), 20)
varEstadoforn = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").item(0).Text)
varCEPform = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/CEP").item(0).Text
varNUMforn = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").item(0).Text
varCODIBGEform = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/cMun").item(0).Text


'Comando de inclusao
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
   Fornecedores!Razao = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xNome").item(0).Text), 35)
   Fornecedores!Fantasia = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/xFant").item(0).Text), 35)
   If tipoEmp = 0 Then Fornecedores!CNPJ = CPF
   If tipoEmp = 1 Then Fornecedores!CNPJ = CNPJ
   Fornecedores!IE = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/IE").item(0).Text
   Fornecedores!Endereco = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xLgr").item(0).Text), 40) & ", " & Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/nro").item(0).Text), 8)
   Fornecedores!bairro = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xBairro").item(0).Text), 20)
   Fornecedores!Cidade = Left(UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/xMun").item(0).Text), 20)
   Fornecedores!Estado = UCase(XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/UF").item(0).Text)
   Fornecedores!CEP = XMLDOC.getElementsByTagName("nfeProc/NFe/infNFe/emit/enderEmit/CEP").item(0).Text
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
    MsgBox "Produto ja adicionado no estoque fisico!", vbInformation, "Aviso do Sistema"
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
      ShowMsg "DATA INVaLIDA!" & vbCrLf & "A data digitada esta incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
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
      ShowMsg "DATA INVaLIDA!" & vbCrLf & "A data digitada esta incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
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
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status " & _
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
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitaaao' END) END) END) AS Status " & _
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
    cboDestOperacao.Text = "1 - Operaaao Interna"
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
    cboIndicadorPagamento.Text = "0 - Pagamento a vista"
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

'CALCULAR PREaO - VAREJO A VISTA
Dim varMargemVV As Currency
Dim varValorVV As Currency

If txtMargemVV.Text = "" Then Exit Sub

varMargemVV = Left$(txtMargemVV.Text, Len(txtMargemVV.Text) - 1)

varValorVV = (varValorCusto * varMargemVV) / 100
varValorVV = varValorCusto + varValorVV
txtValorVV.Text = Format(varValorVV, ocMONEY)

'CALCULAR PREaO - VAREJO A PRAZO
Dim varMargemVP As Currency
Dim varValorVP As Currency

If txtMargemVP.Text = "" Then Exit Sub

varMargemVP = Left$(txtMargemVP.Text, Len(txtMargemVP.Text) - 1)

varValorVP = (varValorCusto * varMargemVP) / 100
varValorVP = varValorCusto + varValorVP
txtValorVP.Text = Format(varValorVP, ocMONEY)

'CALCULAR PREaO - ATACADO A VISTA
Dim varMargemAV As Currency
Dim varValorAV As Currency

If txtMargemAV.Text = "" Then Exit Sub

varMargemAV = Left$(txtMargemAV.Text, Len(txtMargemAV.Text) - 1)

varValorAV = (varValorCusto * varMargemAV) / 100
varValorAV = varValorCusto + varValorAV
txtValorAV.Text = Format(varValorAV, ocMONEY)

'CALCULAR PREaO - ATACADO A PRAZO
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
   MostraStatus_F9$ = "Denegada"  'Deve retornar uma expressao caractere
 ElseIf TbEntrada("Enviada") And Not TbEntrada("Cancelada") Then
   MostraStatus.ForeColor = vbBlue
   MostraStatus_F9$ = "Transmitida/Autorizada"  'Deve retornar uma expressao caractere
 ElseIf TbEntrada("Cancelada") Then
   MostraStatus.ForeColor = vbRed
   MostraStatus_F9$ = "Cancelada"  'Deve retornar uma expressao caractere
 Else
   MostraStatus.ForeColor = vbBlack
   MostraStatus_F9$ = "Em Digitaaao"  'Deve retornar uma expressao caractere
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
   .TextMatrix(0, 1) = "CaD"
   .TextMatrix(0, 2) = "NUM NF"
   .TextMatrix(0, 3) = "EMISSaO"
   .TextMatrix(0, 4) = "FORNECEDOR"
   .TextMatrix(0, 5) = "VALOR"
   .TextMatrix(0, 6) = "STATUS"
   .TextMatrix(0, 7) = "SITUAaaO"
   .TextMatrix(0, 8) = "PENDaNCIAS"
   
   'colocar os cabeaalho em negrito
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
      .TextMatrix(0, 1) = "CaDIGO"
      .TextMatrix(0, 2) = "DESCRIaaO"
      .TextMatrix(0, 3) = "CST"
      .TextMatrix(0, 4) = "UND"
      .TextMatrix(0, 5) = "QTDE"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "SUBTOTAL"
      
      'colocar os cabeaalho em negrito
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
   .TextMatrix(0, 1) = "CaD"
   .TextMatrix(0, 2) = "NUM NF"
   .TextMatrix(0, 3) = "EMISSaO"
   .TextMatrix(0, 4) = "FORNECEDOR"
   .TextMatrix(0, 5) = "VALOR"
   .TextMatrix(0, 6) = "STATUS"
   .TextMatrix(0, 7) = "SITUAaaO"
   .TextMatrix(0, 8) = "PENDaNCIAS"
   
   'colocar os cabeaalho em negrito
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


