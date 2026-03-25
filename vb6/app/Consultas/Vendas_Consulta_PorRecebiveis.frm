VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Vendas_Consulta_PorRecebiveis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE RECEBÍVEIS"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   Icon            =   "Vendas_Consulta_PorRecebiveis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   13215
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdExibirPedidos 
      Height          =   375
      Left            =   60
      TabIndex        =   37
      Top             =   7980
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "EXIBIR VENDAS"
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
      MICON           =   "Vendas_Consulta_PorRecebiveis.frx":23D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   60
      ScaleHeight     =   1905
      ScaleWidth      =   13065
      TabIndex        =   12
      ToolTipText     =   "Imprimir"
      Top             =   780
      Width           =   13095
      Begin VB.ComboBox cboTipoPgto 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1500
         Width           =   2595
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
         Height          =   1755
         Left            =   5400
         TabIndex        =   13
         Top             =   60
         Width           =   7635
         Begin VB.Frame frmFiltro1 
            Height          =   1485
            Left            =   60
            TabIndex        =   14
            Top             =   180
            Width           =   7455
            Begin ChamaleonBtn.chameleonButton cmdCalendario2 
               Height          =   315
               Left            =   2700
               TabIndex        =   43
               Tag             =   "Calendario"
               Top             =   420
               Visible         =   0   'False
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
               MICON           =   "Vendas_Consulta_PorRecebiveis.frx":23EE
               PICN            =   "Vendas_Consulta_PorRecebiveis.frx":240A
               PICH            =   "Vendas_Consulta_PorRecebiveis.frx":475D
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdCalendario1 
               Height          =   315
               Left            =   1080
               TabIndex        =   42
               Tag             =   "Calendario"
               Top             =   420
               Visible         =   0   'False
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
               MICON           =   "Vendas_Consulta_PorRecebiveis.frx":6AB0
               PICN            =   "Vendas_Consulta_PorRecebiveis.frx":6ACC
               PICH            =   "Vendas_Consulta_PorRecebiveis.frx":8E1F
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.ComboBox cboCliente 
               Height          =   315
               Left            =   120
               TabIndex        =   21
               Top             =   1080
               Visible         =   0   'False
               Width           =   4965
            End
            Begin VB.ComboBox cboVendedor 
               Height          =   315
               Left            =   120
               TabIndex        =   20
               Top             =   1080
               Visible         =   0   'False
               Width           =   4965
            End
            Begin VB.TextBox txtCodigo 
               Height          =   315
               Left            =   120
               TabIndex        =   19
               Top             =   420
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboAno 
               Height          =   315
               Left            =   1500
               Sorted          =   -1  'True
               TabIndex        =   18
               Top             =   420
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboMes 
               Height          =   315
               Left            =   120
               TabIndex        =   17
               Top             =   420
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txtCodCliente 
               Appearance      =   0  'Flat
               Height          =   195
               Left            =   4380
               TabIndex        =   16
               Top             =   780
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtCodFunc 
               Appearance      =   0  'Flat
               Height          =   195
               Left            =   3720
               TabIndex        =   15
               Top             =   780
               Visible         =   0   'False
               Width           =   615
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   315
               Left            =   120
               TabIndex        =   22
               Top             =   420
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
               Left            =   1740
               TabIndex        =   23
               Top             =   420
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdCal1 
               Height          =   315
               Left            =   1320
               TabIndex        =   40
               Tag             =   "Calendario"
               Top             =   420
               Visible         =   0   'False
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
               MICON           =   "Vendas_Consulta_PorRecebiveis.frx":B172
               PICN            =   "Vendas_Consulta_PorRecebiveis.frx":B18E
               PICH            =   "Vendas_Consulta_PorRecebiveis.frx":D4E1
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskData 
               Height          =   315
               Left            =   120
               TabIndex        =   41
               Top             =   420
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdLocalizar 
               Height          =   315
               Left            =   3120
               TabIndex        =   44
               Top             =   420
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
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
               MICON           =   "Vendas_Consulta_PorRecebiveis.frx":F834
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
               Height          =   315
               Left            =   4200
               TabIndex        =   45
               Top             =   420
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
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
               MICON           =   "Vendas_Consulta_PorRecebiveis.frx":F850
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdGerarPDF 
               Height          =   315
               Left            =   5280
               TabIndex        =   56
               Top             =   420
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Gerar PDF"
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
               MICON           =   "Vendas_Consulta_PorRecebiveis.frx":F86C
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblClientes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Clientes:"
               Height          =   195
               Left            =   180
               TabIndex        =   32
               Top             =   840
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label lblVendedor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vendedor(a):"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   840
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label lblAte 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "até"
               Height          =   195
               Left            =   1440
               TabIndex        =   30
               Top             =   480
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.Label lblFim 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data final:"
               Height          =   195
               Left            =   1740
               TabIndex        =   29
               Top             =   180
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.Label lblInicio 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data inicial:"
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   180
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.Label lblCodigo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pedido:"
               Height          =   195
               Left            =   120
               TabIndex        =   27
               Top             =   180
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label lblAno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   1500
               TabIndex        =   26
               Top             =   180
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label lblMes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   195
               Left            =   120
               TabIndex        =   25
               Top             =   180
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label lblData 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data::"
               Height          =   195
               Left            =   120
               TabIndex        =   24
               Top             =   180
               Visible         =   0   'False
               Width           =   435
            End
         End
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   2595
      End
      Begin VB.ComboBox cboCriterioPrinc 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   900
         Width           =   2595
      End
      Begin VB.ComboBox cboIndice 
         Height          =   315
         Left            =   2760
         TabIndex        =   1
         Top             =   300
         Width           =   2595
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Pgto"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Classificação"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   60
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Criterio"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Organizar por:"
         Height          =   195
         Left            =   2760
         TabIndex        =   33
         Top             =   60
         Width           =   990
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   5100
      Picture         =   "Vendas_Consulta_PorRecebiveis.frx":F888
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   7
      Top             =   4860
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   13065
      TabIndex        =   3
      Top             =   60
      Width           =   13095
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTA DE RECEBÍVEIS"
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
         Left            =   1440
         TabIndex        =   5
         Top             =   180
         Width           =   4185
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   240
         Picture         =   "Vendas_Consulta_PorRecebiveis.frx":108C0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   11
      Top             =   9375
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18971
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "22:21"
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
      Height          =   5175
      Left            =   60
      TabIndex        =   38
      Top             =   2760
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   9128
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirParcelas 
      Height          =   375
      Left            =   1800
      TabIndex        =   39
      Top             =   7980
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "EXIBIR PARCELAS"
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
      MICON           =   "Vendas_Consulta_PorRecebiveis.frx":17106
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblQtdaHaver 
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
      Left            =   10500
      TabIndex        =   55
      Top             =   8940
      Width           =   915
   End
   Begin VB.Label lblTotalHaver 
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
      Left            =   11460
      TabIndex        =   54
      Top             =   8940
      Width           =   1635
   End
   Begin VB.Label lblQtdaParcela 
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
      Left            =   10500
      TabIndex        =   53
      Top             =   8640
      Width           =   915
   End
   Begin VB.Label lblTotalParcela 
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
      Left            =   11460
      TabIndex        =   52
      Top             =   8640
      Width           =   1635
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HAVERES:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9720
      TabIndex        =   51
      Top             =   8940
      Width           =   720
   End
   Begin VB.Label lblTotalVenda 
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
      Left            =   11460
      TabIndex        =   50
      Top             =   8340
      Width           =   1635
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENDAS:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9810
      TabIndex        =   49
      Top             =   8340
      Width           =   630
   End
   Begin VB.Label lblQtdaVenda 
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
      Left            =   10500
      TabIndex        =   48
      Top             =   8340
      Width           =   915
   End
   Begin VB.Label lblSubtotalBruto 
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
      Left            =   11460
      TabIndex        =   47
      Top             =   8040
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARCELAS:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9615
      TabIndex        =   46
      Top             =   8640
      Width           =   825
   End
   Begin VB.Label lblEntrada 
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
      Left            =   6180
      TabIndex        =   10
      Top             =   8520
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label lblQtda 
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
      Left            =   10500
      TabIndex        =   9
      Top             =   8040
      Width           =   915
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GERAL:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9900
      TabIndex        =   8
      Top             =   8040
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1335
      Left            =   9480
      Top             =   7980
      Width           =   3675
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   8940
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Vendas_Consulta_PorRecebiveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper
Private printSQL As String

Dim posX As Single

Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer

Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long
Private Sub Limpar_Grid_Venda()
   Dim i As Integer

picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 8
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 4300
      .ColWidth(4) = 1000
      .ColWidth(5) = 1100
      .ColWidth(6) = 1220
      .ColWidth(7) = 0
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "NOME DO CLIENTE"
      .TextMatrix(0, 4) = "VALOR"
      .TextMatrix(0, 5) = "FORMA"
      .TextMatrix(0, 6) = "TIPO"
      .TextMatrix(0, 7) = "TIPO"
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      i = 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 1
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 4
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      Grid.Redraw = True
   End With
   
   lblQtda.Caption = Format(0, "000")
   lblQtdaVenda.Caption = Format(0, "000")
   lblQtdaParcela.Caption = Format(0, "000")
   lblQtdaHaver.Caption = Format(0, "000")
   lblSubtotalBruto.Caption = FormatNumber(0, 2)
   lblTotalVenda.Caption = FormatNumber(0, 2)
   lblTotalParcela.Caption = FormatNumber(0, 2)
   lblTotalHaver.Caption = FormatNumber(0, 2)
picAguarde.Visible = False
End Sub

Private Sub LimparObjetos_Consulta()
cboMes.Text = ""
cboAno.Text = ""
cboVendedor.Text = ""
txtCodigo.Text = ""
cboCliente.Text = ""
mskFim.Mask = ""
mskFim.Text = ""
mskInicio.Mask = ""
mskInicio.Text = ""
txtCodFunc.Text = ""
txtCodCliente.Text = ""
End Sub

Private Sub PreencherPrincipal()
cboCriterioPrinc.Clear

'cboCriterioPrinc.AddItem "TODOS"
cboCriterioPrinc.AddItem "PERIODO"
cboCriterioPrinc.AddItem "MENSAL"
End Sub

Private Sub PreencherTipoPgto()
cboTipoPgto.Clear
cboTipoPgto.AddItem "TODOS"
cboTipoPgto.AddItem "DINHEIRO"
cboTipoPgto.AddItem "PIX"
cboTipoPgto.AddItem "CARTÃO DÉBITO"
cboTipoPgto.AddItem "CARTÃO CRÉDITO"
cboTipoPgto.AddItem "TRANSFERÊNCIA"
cboTipoPgto.AddItem "DEPOSITO"
cboTipoPgto.AddItem "FINANCEIRA"
cboTipoPgto.AddItem "CHEQUE"
cboTipoPgto.AddItem "BOLETO"
cboTipoPgto.AddItem "PROMISSÓRIA"
End Sub

Private Sub PreencherIndice()
cboIndice.Clear
'cboIndice.AddItem "PEDIDO"
cboIndice.AddItem "DATA"
'cboIndice.AddItem "POR NOME"
'cboIndice.AddItem "FORMA PGTO"
'cboIndice.AddItem "VALOR"
End Sub



Private Sub PreencherTipoConsulta()
cboTipo.Clear
cboTipo.AddItem "TODOS"
cboTipo.AddItem "VENDAS"
cboTipo.AddItem "PARCELAS"
cboTipo.AddItem "HAVERES"
End Sub

Private Sub cboAno_GotFocus()
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

Private Sub cboAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub cboCliente_Click()
   On Error GoTo TrataErro
   
   If cboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
   If cboCliente.ListIndex = -1 Then txtCodCliente.Text = "": Exit Sub
   txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub CboCliente_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboCliente.Clear
   
   sSQL = "SELECT * FROM cliente ORDER BY nome;"
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

Private Sub CboCliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub CboCliente_LostFocus()
   cboCliente_Click
End Sub

Private Sub cboCriterioPrinc_Change()
   LimparObjetos_Consulta
   
   If cboCriterioPrinc.Text = "TODOS" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      
   ElseIf cboCriterioPrinc.Text = "VENDEDOR" Then
      lblVendedor.Visible = True
      cboVendedor.Visible = True
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      cboVendedor.SetFocus
      
   ElseIf cboCriterioPrinc.Text = "CLIENTE" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario2.Visible = False
      cmdCalendario1.Visible = False
      
      lblClientes.Visible = True
      cboCliente.Visible = True
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      cboCliente.SetFocus
      
   ElseIf cboCriterioPrinc.Text = "PERIODO" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = True
      mskInicio.Visible = True
      lblFim.Visible = True
      mskFim.Visible = True
      lblAte.Visible = True
      cmdCalendario1.Visible = True
      cmdCalendario2.Visible = True
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      mskInicio.SetFocus
      
   ElseIf cboCriterioPrinc.Text = "CÓDIGO" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = True
      txtCodigo.Visible = True
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      txtCodigo.SetFocus
      
   ElseIf cboCriterioPrinc.Text = "MENSAL" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = True
      cboMes.Visible = True
      lblAno.Visible = True
      cboAno.Visible = True
      
      If cboMes.Visible = True Then cboMes.SetFocus
      
   ElseIf cboCriterioPrinc.Text = "CATEGORIA" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
   ElseIf cboCriterioPrinc.Text = "ESPECIFICO" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False

   ElseIf cboCriterioPrinc.Text = "ESPECIFICO/MENSAL" Then
      lblVendedor.Visible = False
      cboVendedor.Visible = False
      
      lblInicio.Visible = False
      mskInicio.Visible = False
      lblFim.Visible = False
      mskFim.Visible = False
      lblAte.Visible = False
      cmdCalendario1.Visible = False
      cmdCalendario2.Visible = False
      
      lblClientes.Visible = False
      cboCliente.Visible = False
      
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      
      lblMes.Visible = True
      cboMes.Visible = True
      lblAno.Visible = True
      cboAno.Visible = True
   Else
      Exit Sub
   End If
End Sub

Private Sub cboCriterioPrinc_Click()
   cboCriterioPrinc_Change
End Sub

Private Sub cboCriterioPrinc_GotFocus()

moCombo.AttachTo cboCriterioPrinc
End Sub

Private Sub cboCriterioPrinc_Validate(Cancel As Boolean)
If cboCriterioPrinc.Text = "ESPECIFICO/MENSAL" Then
   lblMes.Visible = True
   cboMes.Visible = True
   lblAno.Visible = True
   cboAno.Visible = True
End If
End Sub





Private Sub cboIndice_GotFocus()
moCombo.AttachTo cboIndice
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

Private Sub cboMes_LostFocus()
   cboAno.SetFocus
End Sub


Private Sub cboTipo_GotFocus()
moCombo.AttachTo cboTipo
End Sub

Private Sub cboTipoPgto_Click()
   'cboTipoPgto_Change
End Sub

Private Sub cboVendedor_Click()
   cboVendedor_LostFocus
End Sub

Private Sub cboVendedor_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboVendedor.Clear
   
   sSQL = "SELECT codigo, nome, cargo FROM funcionario ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboVendedor.AddItem r("nome")
      cboVendedor.ItemData(cboVendedor.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboVendedor
End Sub

Private Sub cboVendedor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdLocalizar_Click
End Sub

Private Sub cboVendedor_LostFocus()
   On Error GoTo TrataErro
   
   If cboVendedor.Text = "" Then txtCodFunc.Text = "": Exit Sub
   If cboVendedor.ListIndex = -1 Then txtCodFunc.Text = "": Exit Sub
   txtCodFunc = cboVendedor.ItemData(cboVendedor.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
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

mskData = Format(varData, "dd/mm/yyyy")   'Exibe a data no campo
End Sub
Private Sub cmdCalendario1_Click()
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
  
   mskInicio = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdCalendario2_Click()
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
  
   mskFim = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub


Private Sub cmdExibirParcelas_Click()
If Grid.Col = 0 Then Exit Sub
   If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
         Vendas_Consulta_Geral_Parcelas.loadInformacoes (Grid.TextMatrix(Grid.Row, 1))
         Vendas_Consulta_Geral_Parcelas.Show 1
   End If
End Sub

Private Sub cmdExibirPedidos_Click()
If Grid.Col = 0 Then Exit Sub
If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
   If Grid.Col = 1 Then
      If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub
      Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), " VENDAS "
      Parcelas_Consulta_Produtos.Show 1
   End If
End If
End Sub

Private Sub cmdGerarPDF_Click()
Dim r As ADODB.Recordset

'colocar o nome da maquina na barra de status
'Dim var_Impressora As String
'Dim oIni As Ini

'Set oIni = New Ini
'oIni.Arquivo = appPathApp & "config.ini"
'var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
'Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Cons_Recebiveis.Relatorio.Recordset = r
  
REL_Cons_Recebiveis.lblTitulo.Caption = "RELATÓRIO DE RECEBÍVEIS"

REL_Cons_Recebiveis.rfForma.Caption = cboTipoPgto.Text

If cboCriterioPrinc.Text = "PERIODO" Then
   REL_Cons_Recebiveis.rfCons1.Caption = "Intervalo de " & mskInicio.Text & " à " & mskFim.Text
ElseIf cboCriterioPrinc.Text = "MENSAL" Then
   REL_Cons_Recebiveis.rfCons1.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
Else
   REL_Cons_Recebiveis.rfCons1.Caption = "TODOS"
End If

REL_Cons_Recebiveis.dfQuant.Caption = lblQtda.Caption
REL_Cons_Recebiveis.dfQuantVendas.Caption = lblQtdaVenda.Caption
REL_Cons_Recebiveis.dfQuantParcelas.Caption = lblQtdaParcela.Caption
REL_Cons_Recebiveis.dfQuantHaveres.Caption = lblQtdaHaver.Caption
REL_Cons_Recebiveis.dfSubtotalBruto.Caption = lblSubtotalBruto.Caption
REL_Cons_Recebiveis.dfTotalVendas.Caption = lblTotalVenda.Caption
REL_Cons_Recebiveis.dfTotalParcelas.Caption = lblTotalParcela.Caption
REL_Cons_Recebiveis.dfTotalHaveres.Caption = lblTotalHaver.Caption
REL_Cons_Recebiveis.Relatorio.NomeImpressora = "IMPRESSORA PDF"
REL_Cons_Recebiveis.Relatorio.Ativar
Unload REL_Cons_Recebiveis

Me.Show 1

End Sub

Private Sub cmdImprimir_Click()
Dim r As ADODB.Recordset

'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Cons_Recebiveis.Relatorio.Recordset = r
  
REL_Cons_Recebiveis.lblTitulo.Caption = "RELATÓRIO DE RECEBÍVEIS"

REL_Cons_Recebiveis.rfForma.Caption = cboTipoPgto.Text

If cboCriterioPrinc.Text = "PERIODO" Then
   REL_Cons_Recebiveis.rfCons1.Caption = "Intervalo de " & mskInicio.Text & " à " & mskFim.Text
ElseIf cboCriterioPrinc.Text = "MENSAL" Then
   REL_Cons_Recebiveis.rfCons1.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
Else
   REL_Cons_Recebiveis.rfCons1.Caption = "TODOS"
End If

REL_Cons_Recebiveis.dfQuant.Caption = lblQtda.Caption
REL_Cons_Recebiveis.dfQuantVendas.Caption = lblQtdaVenda.Caption
REL_Cons_Recebiveis.dfQuantParcelas.Caption = lblQtdaParcela.Caption
REL_Cons_Recebiveis.dfQuantHaveres.Caption = lblQtdaHaver.Caption
REL_Cons_Recebiveis.dfSubtotalBruto.Caption = lblSubtotalBruto.Caption
REL_Cons_Recebiveis.dfTotalVendas.Caption = lblTotalVenda.Caption
REL_Cons_Recebiveis.dfTotalParcelas.Caption = lblTotalParcela.Caption
REL_Cons_Recebiveis.dfTotalHaveres.Caption = lblTotalHaver.Caption

REL_Cons_Recebiveis.Relatorio.Ativar
Unload REL_Cons_Recebiveis

Me.Show 1
End Sub

Public Sub cmdLocalizar_Click()
totalRegistros = "0"
Dim vTABELAS As String
Dim SQLCount As String
Dim SQLWhere As String
Dim vWhere As String
Dim vPegarCliente As String

'INDICE
Dim INDICE As String
If cboIndice.Text = "DATA" Then
   INDICE = "vDataPgto;"
ElseIf cboIndice.Text = "POR NOME" Then
   INDICE = "nome;"
ElseIf cboIndice.Text = "FORMA PGTO" Then
   INDICE = "tipo_pagamento;"
ElseIf cboIndice.Text = "VALOR" Then
   INDICE = "pedidos_1.total;"
ElseIf cboIndice.Text = "PEDIDO" Then
   INDICE = "pedidos_1.cod_pedido;"
Else
   INDICE = "pedidos_1.cod_pedido;"
End If

'TIPO DE PAGAMENTO (PARCELAS)
Dim vTipoPgtoParcelas As String
'If cboFormaPgto.Text = "À VISTA" Then
    If cboTipoPgto.Text = "TODOS" Then
        vTipoPgtoParcelas = " <> 'BOSTA'"
    ElseIf cboTipoPgto.Text = "DINHEIRO" Then
        vTipoPgtoParcelas = " = 'DINHEIRO'"
    ElseIf cboTipoPgto.Text = "CARTÃO DÉBITO" Then
        vTipoPgtoParcelas = " = 'CARTAO'"
        'vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'CARTAO') and (parcelas.TIPO_CARTAO = 'D')"""
    ElseIf cboTipoPgto.Text = "CARTÃO CRÉDITO" Then
        vTipoPgtoParcelas = " = 'CARTAO'"
        'vTipoPgtoParcelas = " AND (parcelas.FORMA_PGTO = 'CARTAO') and (parcelas.TIPO_CARTAO = 'C')"""
    ElseIf cboTipoPgto.Text = "TRANSFERÊNCIA" Then
        vTipoPgtoParcelas = " = 'TRANSFERENCIA'"
    ElseIf cboTipoPgto.Text = "DEPOSITO" Then
        vTipoPgtoParcelas = " = 'DEPOSITO'"
    ElseIf cboTipoPgto.Text = "FINANCEIRA" Then
        vTipoPgtoParcelas = " = 'FINANCEIRA'"
    ElseIf cboTipoPgto.Text = "PROMISSÓRIA" Then
        vTipoPgtoParcelas = " = 'PROMISSORIA'"
    ElseIf cboTipoPgto.Text = "CHEQUE" Then
        vTipoPgtoParcelas = " = 'CHEQUE'"
    ElseIf cboTipoPgto.Text = "BOLETO" Then
        vTipoPgtoParcelas = " = 'BOLETO'"
    ElseIf cboTipoPgto.Text = "PIX" Then
        vTipoPgtoParcelas = " = 'PIX'"
    End If
    
Dim vTipoCartao As String
'If cboFormaPgto.Text = "À VISTA" Then
    If cboTipoPgto.Text = "CARTÃO DÉBITO" Or cboTipoPgto.Text = "CARTÃO CRÉDITO" Then
        If cboTipoPgto.Text = "CARTÃO DÉBITO" Then
            vTipoCartao = "AND (parcelas.TIPO_CARTAO = 'D')"
            'vTipoCartao = " = 'D'"
        Else
            vTipoCartao = "AND (parcelas.TIPO_CARTAO = 'C')"
            'vTipoCartao = " = 'C'"
        End If
    Else
        If cboTipoPgto.Text = "TODOS" Then
            vTipoCartao = " "
        Else
            vTipoCartao = "AND (parcelas.TIPO_CARTAO IS NULL)"
            'vTipoCartao = " IS NULL"
        End If
    End If

Dim vTipoCartaoHaver As String
'If cboFormaPgto.Text = "À VISTA" Then
    If cboTipoPgto.Text = "CARTÃO DÉBITO" Or cboTipoPgto.Text = "CARTÃO CRÉDITO" Then
        If cboTipoPgto.Text = "CARTÃO DÉBITO" Then
            vTipoCartaoHaver = "AND (parcelas_haver.TIPO_CARTAO = 'D')"
            'vTipoCartao = " = 'D'"
        Else
            vTipoCartaoHaver = "AND (parcelas_haver.TIPO_CARTAO = 'C')"
            'vTipoCartao = " = 'C'"
        End If
    Else
        If cboTipoPgto.Text = "TODOS" Then
            vTipoCartaoHaver = " "
        Else
            vTipoCartaoHaver = "AND (parcelas_haver.TIPO_CARTAO IS NULL)"
            'vTipoCartao = " IS NULL"
        End If
    End If
    
'TIPO DE CONSULTA
Dim varTipoConsulta As String
If cboTipo.Text = "TODOS" Then
   varTipoConsulta = "VENDA"
ElseIf cboTipo.Text = "VENDAS" Then
   varTipoConsulta = "VENDA"
ElseIf cboTipo.Text = "PARCELAS" Then
   varTipoConsulta = "PARCELA"
ElseIf cboTipo.Text = "HAVERES" Then
   varTipoConsulta = "HAVER"
End If

If cboCriterioPrinc.Text = "" Then Exit Sub
   
   'PERIODO
    If cboCriterioPrinc.Text = "PERIODO" Then
        If Not IsDate(mskInicio) Or Not IsDate(mskFim) Then Limpar_Grid_Venda: Exit Sub
                If cboTipo.Text = "TODOS" Then
                sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas.PAGAMENTO AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas.FORMA_PGTO END AS pFormaPgto, " & _
                    "(parcelas.TIPO) AS vTipoPgto,  " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome  " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO  " & _
                    "WHERE (parcelas.PAGAMENTO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (parcelas.PAGAMENTO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (parcelas.FORMA_PGTO " & vTipoPgtoParcelas & ") " & vTipoCartao & " " & _
                    "UNION ALL "
                sSQL = sSQL & "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas_haver.HAVER AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas_haver.VALOR_HAVER AS pValorPgto, parcelas_haver.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas_haver.FORMA_PGTO END AS pFormaPgto, " & _
                    "'HAVER' AS vTipoPgto, " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN parcelas_haver ON parcelas.CODIGO = parcelas_haver.COD_PARCELA " & _
                    "WHERE (parcelas_haver.HAVER >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (parcelas_haver.HAVER <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (parcelas_haver.FORMA_PGTO " & vTipoPgtoParcelas & ") " & vTipoCartaoHaver & " " & _
                    "ORDER BY " & INDICE
            ElseIf cboTipo.Text = "VENDAS" Then
                sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas.PAGAMENTO AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas.FORMA_PGTO END AS pFormaPgto, " & _
                    "(parcelas.TIPO) AS vTipoPgto,  " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome  " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO  " & _
                    "WHERE (parcelas.PAGAMENTO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (parcelas.PAGAMENTO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (parcelas.FORMA_PGTO " & vTipoPgtoParcelas & ") " & vTipoCartao & " AND parcelas.TIPO = '" & varTipoConsulta & "' " & _
                    "ORDER BY " & INDICE
            ElseIf cboTipo.Text = "PARCELAS" Then
                sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas.PAGAMENTO AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas.FORMA_PGTO END AS pFormaPgto, " & _
                    "(parcelas.TIPO) AS vTipoPgto,  " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome  " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO  " & _
                    "WHERE (parcelas.PAGAMENTO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (parcelas.PAGAMENTO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (parcelas.FORMA_PGTO " & vTipoPgtoParcelas & ") " & vTipoCartao & " AND parcelas.TIPO = '" & varTipoConsulta & "' " & _
                    "ORDER BY " & INDICE
            ElseIf cboTipo.Text = "HAVERES" Then
                sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas_haver.HAVER AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas_haver.VALOR_HAVER AS pValorPgto, parcelas_haver.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas_haver.FORMA_PGTO END AS pFormaPgto, " & _
                    "'HAVER' AS vTipoPgto, " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN parcelas_haver ON parcelas.CODIGO = parcelas_haver.COD_PARCELA " & _
                    "WHERE (parcelas_haver.HAVER >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (parcelas_haver.HAVER <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (parcelas_haver.FORMA_PGTO " & vTipoPgtoParcelas & ") " & vTipoCartao & " " & _
                    "ORDER BY " & INDICE
            End If


        ''sSQL = sSQL & vTABELAS & " Where (pedidos_1.data_compra >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_1.data_compra <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & "  "
        ''SQLWhere = " Where (pedidos_1.data_compra >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_1.data_compra <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') and cancelado = '0' " & vTipoPgtoParcelas & "  "
        
        ''sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE, parcelas.FORMA_PGTO AS pFormaPgto, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
            "(CASE pedidos_1.TIPO_PAGAMENTO WHEN 'À Prazo' THEN pedidos_1.PAGAMENTO ELSE parcelas.FORMA_PGTO END) AS vTipoPgto, " & _
            "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO WHERE (pedidos_1.data_compra >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_1.data_compra <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = '0') ) AS Nome " & _
            "FROM pedidos AS pedidos_1 INNER JOIN  parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO " & _
            "WHERE (pedidos_1.data_compra >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_1.data_compra <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (pedidos_1.tipo_pedido = '" & varTipoConsulta & "') AND (pedidos_1.CANCELADO = 0)  " & _
            "ORDER BY var_codped"
            
           ' If cboTipo.Text = "TODOS" Then
           '     sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas.PAGAMENTO AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas.FORMA_PGTO END AS pFormaPgto, " & _
                    "(parcelas.TIPO) AS vTipoPgto,  " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome  " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO  " & _
                    "WHERE (parcelas.PAGAMENTO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (parcelas.PAGAMENTO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (parcelas.FORMA_PGTO " & vTipoPgtoParcelas & ") AND (parcelas.TIPO_CARTAO " & vTipoCartao & ") " & _
                    "UNION ALL "
           '     sSQL = sSQL & "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas_haver.HAVER AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas_haver.VALOR_HAVER AS pValorPgto, parcelas_haver.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas_haver.FORMA_PGTO END AS pFormaPgto, " & _
                    "'HAVER' AS vTipoPgto, " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN parcelas_haver ON parcelas.CODIGO = parcelas_haver.COD_PARCELA " & _
                    "WHERE (parcelas_haver.HAVER >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (parcelas_haver.HAVER <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (parcelas_haver.FORMA_PGTO " & vTipoPgtoParcelas & ") AND (parcelas_haver.TIPO_CARTAO " & vTipoCartao & ") " & _
                    "ORDER BY " & INDICE
           ' ElseIf cboTipo.Text = "VENDAS" Then
           '     sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas.PAGAMENTO AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas.FORMA_PGTO END AS pFormaPgto, " & _
                    "(parcelas.TIPO) AS vTipoPgto,  " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome  " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO  " & _
                    "WHERE (parcelas.PAGAMENTO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (parcelas.PAGAMENTO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (parcelas.FORMA_PGTO " & vTipoPgtoParcelas & ") AND (parcelas.TIPO_CARTAO " & vTipoCartao & ") AND parcelas.TIPO = '" & varTipoConsulta & "' " & _
                    "ORDER BY " & INDICE
           ' ElseIf cboTipo.Text = "PARCELAS" Then
           '     sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas.PAGAMENTO AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas.FORMA_PGTO END AS pFormaPgto, " & _
                    "(parcelas.TIPO) AS vTipoPgto,  " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome  " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO  " & _
                    "WHERE (parcelas.PAGAMENTO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (parcelas.PAGAMENTO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (parcelas.FORMA_PGTO " & vTipoPgtoParcelas & ") AND (parcelas.TIPO_CARTAO " & vTipoCartao & ") AND parcelas.TIPO = '" & varTipoConsulta & "' " & _
                    "ORDER BY " & INDICE
           ' ElseIf cboTipo.Text = "HAVERES" Then
           '     sSQL = sSQL & "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas_haver.HAVER AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas_haver.VALOR_HAVER AS pValorPgto, parcelas_haver.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas_haver.FORMA_PGTO END AS pFormaPgto, " & _
                    "'HAVER' AS vTipoPgto, " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN parcelas_haver ON parcelas.CODIGO = parcelas_haver.COD_PARCELA " & _
                    "WHERE (parcelas_haver.HAVER >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (parcelas_haver.HAVER <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) AND (parcelas_haver.FORMA_PGTO " & vTipoPgtoParcelas & ") AND (parcelas_haver.TIPO_CARTAO " & vTipoCartao & ") " & _
                    "ORDER BY " & INDICE
           ' End If
            
'AND (parcelas.TIPO_CARTAO " & vTipoCartao & ")
'AND (parcelas_haver.TIPO_CARTAO " & vTipoCartao & ")

   'MENSAL
    ElseIf cboCriterioPrinc.Text = "MENSAL" Then
        If cboMes.Text = "" Or cboAno.Text = "" Then Limpar_Grid_Venda: Exit Sub
            If cboTipo.Text = "TODOS" Then
                sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas.PAGAMENTO AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas.FORMA_PGTO END AS pFormaPgto, " & _
                    "(parcelas.TIPO) AS vTipoPgto,  " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome  " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO  " & _
                    "WHERE (MONTH(parcelas.PAGAMENTO) = " & cboMes.ListIndex + 1 & ") AND (YEAR(parcelas.PAGAMENTO) = " & cboAno & ") AND (parcelas.FORMA_PGTO " & vTipoPgtoParcelas & ") " & vTipoCartao & " " & _
                    "UNION ALL "
                sSQL = sSQL & "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas_haver.HAVER AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas_haver.VALOR_HAVER AS pValorPgto, parcelas_haver.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas_haver.FORMA_PGTO END AS pFormaPgto, " & _
                    "'HAVER' AS vTipoPgto, " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN parcelas_haver ON parcelas.CODIGO = parcelas_haver.COD_PARCELA " & _
                    "WHERE (MONTH(parcelas_haver.HAVER) = " & cboMes.ListIndex + 1 & ") AND (YEAR(parcelas_haver.HAVER) = " & cboAno & ") AND (parcelas_haver.FORMA_PGTO " & vTipoPgtoParcelas & ") " & vTipoCartaoHaver & " " & _
                    "ORDER BY " & INDICE
            ElseIf cboTipo.Text = "VENDAS" Then
                sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas.PAGAMENTO AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas.FORMA_PGTO END AS pFormaPgto, " & _
                    "(parcelas.TIPO) AS vTipoPgto,  " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome  " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO  " & _
                    "WHERE (MONTH(parcelas.PAGAMENTO) = " & cboMes.ListIndex + 1 & ") AND (YEAR(parcelas.PAGAMENTO) = " & cboAno & ") AND (parcelas.FORMA_PGTO " & vTipoPgtoParcelas & ") " & vTipoCartao & " AND parcelas.TIPO = '" & varTipoConsulta & "' " & _
                    "ORDER BY " & INDICE
            ElseIf cboTipo.Text = "PARCELAS" Then
                sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas.PAGAMENTO AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas.VALOR_FINAL AS pValorPgto, parcelas.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas.FORMA_PGTO = 'CARTAO' AND parcelas.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas.FORMA_PGTO END AS pFormaPgto, " & _
                    "(parcelas.TIPO) AS vTipoPgto,  " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome  " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO  " & _
                    "WHERE (MONTH(parcelas.PAGAMENTO) = " & cboMes.ListIndex + 1 & ") AND (YEAR(parcelas.PAGAMENTO) = " & cboAno & ") AND (parcelas.FORMA_PGTO " & vTipoPgtoParcelas & ") " & vTipoCartao & " AND parcelas.TIPO = '" & varTipoConsulta & "' " & _
                    "ORDER BY " & INDICE
            ElseIf cboTipo.Text = "HAVERES" Then
                sSQL = "SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, parcelas_haver.HAVER AS vDataPgto, parcelas.DATA, pedidos_1.COD_CLIENTE, parcelas_haver.VALOR_HAVER AS pValorPgto, parcelas_haver.TIPO_CARTAO, " & _
                    "CASE WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'C') THEN 'CARTÃO CRÉDITO' WHEN (parcelas_haver.FORMA_PGTO = 'CARTAO' AND parcelas_haver.TIPO_CARTAO = 'D') THEN 'CARTÃO DÉBITO' ELSE parcelas_haver.FORMA_PGTO END AS pFormaPgto, " & _
                    "'HAVER' AS vTipoPgto, " & _
                    "(SELECT DISTINCT cliente.Nome FROM pedidos INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO) AS Nome " & _
                    "FROM pedidos AS pedidos_1 INNER JOIN parcelas ON pedidos_1.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN parcelas_haver ON parcelas.CODIGO = parcelas_haver.COD_PARCELA " & _
                    "WHERE (MONTH(parcelas_haver.HAVER) = " & cboMes.ListIndex + 1 & ") AND (YEAR(parcelas_haver.HAVER) = " & cboAno & ") AND (parcelas_haver.FORMA_PGTO " & vTipoPgtoParcelas & ") " & vTipoCartao & " " & _
                    "ORDER BY " & INDICE
            End If
   End If
 
'Debug.Print sSQL

Set r = dbData.OpenRecordset(sSQL, totalRegistros)
lblQtda.Caption = Format(totalRegistros, "00")

FormatarGrid_Vendas r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Double
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   
   For i = 0 To var_Grid.rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub Form_Load()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing
   
'limpar o grid

PreencherTipoConsulta
cboTipo.ListIndex = 0

PreencherIndice
cboIndice.ListIndex = 0

PreencherPrincipal
cboCriterioPrinc.ListIndex = 1

PreencherTipoPgto
cboTipoPgto.ListIndex = 0

StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
Set moCombo = New cComboHelper
End Sub

Private Sub SomaFlexMetodos()
On Error GoTo errorhandeler
Dim soma As Currency
Dim QUANT As Integer
Dim i As Integer

soma = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 3) = "VENDA" And IsNumeric(.TextMatrix(i, 5)) Then
         soma = soma + CCur(.TextMatrix(i, 5))
         QUANT = QUANT + 1
      End If
   Next
End With

lblTotalVenda.Caption = FormatNumber(soma, 2)
lblQtdaVenda.Caption = Format(QUANT, "000")

soma = 0
QUANT = 0

With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 3) = "PARCELA" And IsNumeric(.TextMatrix(i, 5)) Then
         soma = soma + CCur(.TextMatrix(i, 5))
         QUANT = QUANT + 1
      End If
   Next
End With

lblTotalParcela.Caption = FormatNumber(soma, 2)
lblQtdaParcela.Caption = Format(QUANT, "000")

soma = 0
QUANT = 0

With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 3) = "HAVER" And IsNumeric(.TextMatrix(i, 5)) Then
         soma = soma + CCur(.TextMatrix(i, 5))
         QUANT = QUANT + 1
      End If
   Next
End With

lblTotalHaver.Caption = FormatNumber(soma, 2)
lblQtdaHaver.Caption = Format(QUANT, "000")

soma = 0
QUANT = 0

errorhandeler:
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   posX = x
   Label3 = posX
   If Label3.Caption > 0 And Label3.Caption < 149 Then Grid.ToolTipText = ""
   If Label3.Caption > 150 And Label3.Caption < 930 Then Grid.ToolTipText = "Dê um duplo-clique para exibir os itens do Pedido."
   If Label3.Caption > 931 And Label3.Caption < 7230 Then Grid.ToolTipText = ""
   If Label3.Caption > 7231 And Label3.Caption < 8355 Then Grid.ToolTipText = "Dê um duplo-clique para exibir a forma de pgto."
   If Label3.Caption > 8356 And Label3.Caption < 9555 Then Grid.ToolTipText = ""
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
mskData.Mask = "##/##/##"
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

Sub FormatarGrid_Vendas(rTabela As ADODB.Recordset)
Dim i As Integer
picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 7
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 900
      .ColWidth(3) = 1300
      .ColWidth(4) = 5500
      .ColWidth(5) = 1000
      .ColWidth(6) = 2300
     
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "METODO"
      .TextMatrix(0, 4) = "NOME DO CLIENTE"
      .TextMatrix(0, 5) = "VALOR"
      .TextMatrix(0, 6) = "FORMA"
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      i = 1
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("var_codped"), "000000")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("vdatapgto"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("vTipoPgto"))
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("nome"))
            .TextMatrix(.rows - 1, 5) = FormatNumber(rTabela("pvalorPgto"), 2)
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("pFormaPgto"))

            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If
      
     
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 1
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 6
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      Grid.Redraw = True
   End With
   
    lblSubtotalBruto.Caption = FormatNumber(SomaGrid(Grid, 5), 2)
    SomaFlexMetodos
picAguarde.Visible = False
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
