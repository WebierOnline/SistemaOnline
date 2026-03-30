VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Caixa_Controle_semOS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIXA"
   ClientHeight    =   9975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13080
   Icon            =   "Caixa_Controle.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmEntradas 
      Caption         =   "FINANCEIRO"
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
      Height          =   4395
      Left            =   8760
      TabIndex        =   51
      Top             =   5280
      Width           =   4215
      Begin VB.TextBox txtTotalOSPrazo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   120
         Top             =   5400
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtQuantOSPrazo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   119
         Top             =   5400
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtQuantAluguelPrazo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   5100
         Width           =   435
      End
      Begin VB.TextBox txtTotalAluguelPrazo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   116
         Top             =   5100
         Width           =   1515
      End
      Begin ChamaleonBtn.chameleonButton cmdAbriFaturamento 
         Height          =   285
         Left            =   3660
         TabIndex        =   96
         Top             =   4500
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "+"
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
         MICON           =   "Caixa_Controle.frx":23D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtQuantDinheiroAluguel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   94
         ToolTipText     =   "DINHEIRO"
         Top             =   1800
         Width           =   435
      End
      Begin VB.TextBox txtTotalDinheiroAluguel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   93
         ToolTipText     =   "DINHEIRO"
         Top             =   1800
         Width           =   1515
      End
      Begin VB.TextBox txtQuantDinheiroOS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   91
         ToolTipText     =   "DINHEIRO"
         Top             =   1500
         Width           =   435
      End
      Begin VB.TextBox txtTotalDinheiroOS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   90
         ToolTipText     =   "DINHEIRO"
         Top             =   1500
         Width           =   1515
      End
      Begin VB.TextBox txtTotalRetiradas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   480
         Width           =   1995
      End
      Begin VB.CheckBox chkTroco 
         Height          =   195
         Left            =   3720
         TabIndex        =   87
         Top             =   180
         Width           =   195
      End
      Begin VB.TextBox txtTotalDinheiro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   73
         ToolTipText     =   "DINHEIRO"
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox txtTotalDinheiroParcelas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   72
         ToolTipText     =   "DINHEIRO"
         Top             =   2100
         Width           =   1515
      End
      Begin VB.TextBox txtTotalDinheiroHaveres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   71
         ToolTipText     =   "DINHEIRO"
         Top             =   2400
         Width           =   1515
      End
      Begin VB.TextBox txtTotalCheque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   3000
         Width           =   1515
      End
      Begin VB.TextBox txtTotalTroco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   180
         Width           =   1995
      End
      Begin VB.TextBox txtSaldoFisico 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   900
         Width           =   1995
      End
      Begin VB.TextBox txtTotalDinheiroSuprimento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   67
         ToolTipText     =   "DINHEIRO"
         Top             =   2700
         Width           =   1515
      End
      Begin VB.TextBox txtSaida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   3300
         Width           =   1515
      End
      Begin VB.TextBox txtTotalCartao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   3900
         Width           =   1515
      End
      Begin VB.TextBox txtTotalAvulso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   4200
         Width           =   1515
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   3600
         Width           =   1995
      End
      Begin VB.TextBox txtTotalVendaPrazo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   4800
         Width           =   1515
      End
      Begin VB.TextBox txtFaturamento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   4500
         Width           =   1995
      End
      Begin VB.TextBox txtQuantVendaPrazo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   4800
         Width           =   435
      End
      Begin VB.TextBox txtQuantAvulso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   4200
         Width           =   435
      End
      Begin VB.TextBox txtQuantCartao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   3900
         Width           =   435
      End
      Begin VB.TextBox txtQuantSaida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   3300
         Width           =   435
      End
      Begin VB.TextBox txtQuantDinheiroSuprimento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   56
         ToolTipText     =   "DINHEIRO"
         Top             =   2700
         Width           =   435
      End
      Begin VB.TextBox txtQuantCheque 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   3000
         Width           =   435
      End
      Begin VB.TextBox txtQuantDinheiroHaveres 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   54
         ToolTipText     =   "DINHEIRO"
         Top             =   2400
         Width           =   435
      End
      Begin VB.TextBox txtQuantDinheiroParcelas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   53
         ToolTipText     =   "DINHEIRO"
         Top             =   2100
         Width           =   435
      End
      Begin VB.TextBox txtQuantDinheiro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   52
         ToolTipText     =   "DINHEIRO"
         Top             =   1200
         Width           =   435
      End
      Begin ChamaleonBtn.chameleonButton cmdAbrirSaldoFisico 
         Height          =   285
         Left            =   3660
         TabIndex        =   97
         Top             =   900
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "+"
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
         MICON           =   "Caixa_Controle.frx":23EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAbrirSaldoGeral 
         Height          =   285
         Left            =   3660
         TabIndex        =   98
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "+"
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
         MICON           =   "Caixa_Controle.frx":240A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblOSFat 
         AutoSize        =   -1  'True
         Caption         =   "OS:"
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
         Left            =   1320
         TabIndex        =   121
         Top             =   5400
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblAluguelFat 
         AutoSize        =   -1  'True
         Caption         =   "ALUGUEL:"
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
         Left            =   720
         TabIndex        =   118
         Top             =   5100
         Width           =   930
      End
      Begin VB.Image img9 
         Height          =   225
         Left            =   3660
         Picture         =   "Caixa_Controle.frx":2426
         Top             =   480
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblAluguel 
         AutoSize        =   -1  'True
         Caption         =   "ALUGUEL*:"
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
         Left            =   600
         TabIndex        =   95
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lblOS 
         AutoSize        =   -1  'True
         Caption         =   "O.S.*:"
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
         TabIndex        =   92
         Top             =   1500
         Width           =   525
      End
      Begin VB.Line Line1 
         BorderStyle     =   6  'Inside Solid
         X1              =   240
         X2              =   4020
         Y1              =   820
         Y2              =   820
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RETIRADAS:"
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
         Left            =   540
         TabIndex        =   89
         Top             =   480
         Width           =   1140
      End
      Begin VB.Image img8 
         Height          =   225
         Left            =   3720
         Picture         =   "Caixa_Controle.frx":28AD
         Top             =   4200
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblVendas 
         AutoSize        =   -1  'True
         Caption         =   "VENDAS*:"
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
         Left            =   750
         TabIndex        =   86
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblParcelas 
         AutoSize        =   -1  'True
         Caption         =   "PARCELAS*:"
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
         Left            =   540
         TabIndex        =   85
         Top             =   2100
         Width           =   1110
      End
      Begin VB.Label lblHaveres 
         AutoSize        =   -1  'True
         Caption         =   "HAVERES*:"
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
         Left            =   630
         TabIndex        =   84
         Top             =   2385
         Width           =   1020
      End
      Begin VB.Label lblCheque 
         AutoSize        =   -1  'True
         Caption         =   "CHEQUES*:"
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
         Left            =   615
         TabIndex        =   83
         Top             =   3000
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TROCO:"
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
         Left            =   930
         TabIndex        =   82
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblSaldoFisico 
         AutoSize        =   -1  'True
         Caption         =   "SALDO FÍSICO:"
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
         Left            =   300
         TabIndex        =   81
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label lblSuprimentos 
         AutoSize        =   -1  'True
         Caption         =   "SUPRIMENTOS*:"
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
         TabIndex        =   80
         Top             =   2700
         Width           =   1500
      End
      Begin VB.Label lblSaidas 
         AutoSize        =   -1  'True
         Caption         =   "SAÍDAS:"
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
         Left            =   915
         TabIndex        =   79
         Top             =   3300
         Width           =   735
      End
      Begin VB.Label lblCartao 
         AutoSize        =   -1  'True
         Caption         =   "CARTÃO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   825
         TabIndex        =   78
         Top             =   3900
         Width           =   825
      End
      Begin VB.Label lblOutros 
         AutoSize        =   -1  'True
         Caption         =   "OUTROS**:"
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
         Left            =   645
         TabIndex        =   77
         Top             =   4200
         Width           =   1005
      End
      Begin VB.Label lblSaldoGeral 
         AutoSize        =   -1  'True
         Caption         =   "SALDO GERAL:"
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
         Left            =   285
         TabIndex        =   76
         Top             =   3600
         Width           =   1365
      End
      Begin VB.Label lblPrazo 
         AutoSize        =   -1  'True
         Caption         =   "À PRAZO:"
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
         Left            =   780
         TabIndex        =   75
         Top             =   4800
         Width           =   870
      End
      Begin VB.Label lblFaturamento 
         AutoSize        =   -1  'True
         Caption         =   "FATURAMENTO:"
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
         TabIndex        =   74
         Top             =   4500
         Width           =   1470
      End
      Begin VB.Image img7 
         Height          =   225
         Left            =   3720
         Picture         =   "Caixa_Controle.frx":2D34
         Top             =   3900
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.Frame frmSenha 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6600
      TabIndex        =   21
      Top             =   6420
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin ChamaleonBtn.chameleonButton cmdSenha 
         Height          =   315
         Left            =   1500
         TabIndex        =   23
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "OK"
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
         MICON           =   "Caixa_Controle.frx":31BB
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
   Begin VB.Frame frmTroco 
      Caption         =   "Troco"
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
      Left            =   5880
      TabIndex        =   8
      Top             =   7620
      Visible         =   0   'False
      Width           =   2835
      Begin VB.TextBox txtTroco 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1395
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarTroco 
         Height          =   375
         Left            =   1620
         TabIndex        =   11
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Salvar"
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
         MICON           =   "Caixa_Controle.frx":31D7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         Caption         =   "Valor"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frmFaturamento 
      Caption         =   "FATURAMENTO"
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
      Height          =   2955
      Left            =   60
      TabIndex        =   30
      Top             =   5640
      Width           =   3915
      Begin VB.TextBox txtF8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   2280
         Width           =   450
      End
      Begin VB.TextBox txtFATTotalAluguel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   2280
         Width           =   1515
      End
      Begin VB.TextBox txtFATTotalServicos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   1980
         Width           =   1515
      End
      Begin VB.TextBox txtF7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   1980
         Width           =   450
      End
      Begin VB.TextBox txtF2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   480
         Width           =   450
      End
      Begin VB.TextBox txtF1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   180
         Width           =   450
      End
      Begin VB.TextBox txtF3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   780
         Width           =   450
      End
      Begin VB.TextBox txtF4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1080
         Width           =   450
      End
      Begin VB.TextBox txtF6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1680
         Width           =   450
      End
      Begin VB.TextBox txtF5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1380
         Width           =   450
      End
      Begin VB.TextBox txtFATTotalPrazo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1380
         Width           =   1515
      End
      Begin VB.TextBox txtFATTotalSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2580
         Width           =   1995
      End
      Begin VB.TextBox txtFATTotalSaidas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1680
         Width           =   1515
      End
      Begin VB.TextBox txtFATTotalSuprimentos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1080
         Width           =   1515
      End
      Begin VB.TextBox txtFATTotalHaveres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   780
         Width           =   1515
      End
      Begin VB.TextBox txtFATTotalVendas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   180
         Width           =   1515
      End
      Begin VB.TextBox txtFATTotalParcelas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   1515
      End
      Begin VB.Image img11 
         Height          =   225
         Left            =   3360
         Picture         =   "Caixa_Controle.frx":31F3
         Top             =   2280
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblFatAluguel 
         AutoSize        =   -1  'True
         Caption         =   "ALUGUEL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   390
         TabIndex        =   115
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label lblFatServicos 
         AutoSize        =   -1  'True
         Caption         =   "SERVIÇOS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   390
         TabIndex        =   101
         Top             =   1980
         Width           =   855
      End
      Begin VB.Image img10 
         Height          =   225
         Left            =   3360
         Picture         =   "Caixa_Controle.frx":367A
         Top             =   1980
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Image img6 
         Height          =   225
         Left            =   3360
         Picture         =   "Caixa_Controle.frx":3B01
         Top             =   1680
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Image img5 
         Height          =   225
         Left            =   3360
         Picture         =   "Caixa_Controle.frx":3F88
         Top             =   1380
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Image img4 
         Height          =   225
         Left            =   3360
         Picture         =   "Caixa_Controle.frx":440F
         Top             =   1080
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Image img3 
         Height          =   225
         Left            =   3360
         Picture         =   "Caixa_Controle.frx":4896
         Top             =   780
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Image img2 
         Height          =   225
         Left            =   3360
         Picture         =   "Caixa_Controle.frx":4D1D
         Top             =   480
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Image img1 
         Height          =   225
         Left            =   3360
         Picture         =   "Caixa_Controle.frx":51A4
         Top             =   180
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "À PRAZO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   44
         Top             =   1380
         Width           =   765
      End
      Begin VB.Label lblSaldo 
         AutoSize        =   -1  'True
         Caption         =   "SALDO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   645
         TabIndex        =   42
         Top             =   2580
         Width           =   600
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "SAÍDAS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   40
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "SUPRIMENTOS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   38
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "HAVERES:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   465
         TabIndex        =   35
         Top             =   780
         Width           =   780
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "VENDAS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   555
         TabIndex        =   34
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "PARCELAS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   330
         TabIndex        =   32
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.Frame frmMaquina 
      Caption         =   "Caixa"
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
      Left            =   5880
      TabIndex        =   13
      Top             =   7620
      Visible         =   0   'False
      Width           =   2835
      Begin VB.ComboBox cboMaquina 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   540
         Width           =   1815
      End
      Begin ChamaleonBtn.chameleonButton cmdMaqOK 
         Height          =   315
         Left            =   1980
         TabIndex        =   15
         Top             =   540
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "OK"
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
         MICON           =   "Caixa_Controle.frx":562B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Caption         =   "Caixa"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame frmSetor 
      Caption         =   "Setor"
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
      Left            =   5880
      TabIndex        =   24
      Top             =   7620
      Visible         =   0   'False
      Width           =   2835
      Begin VB.ComboBox cboSetor 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   540
         Width           =   1815
      End
      Begin ChamaleonBtn.chameleonButton cmdOKSetor 
         Height          =   315
         Left            =   1980
         TabIndex        =   26
         Top             =   540
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "OK"
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
         MICON           =   "Caixa_Controle.frx":5647
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Setor"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "Data"
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
      Left            =   6540
      TabIndex        =   17
      Top             =   7680
      Visible         =   0   'False
      Width           =   2175
      Begin MSMask.MaskEdBox mskData 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton cmdCal1 
         Height          =   315
         Left            =   1200
         TabIndex        =   29
         Tag             =   "Calendario"
         Top             =   540
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
         MICON           =   "Caixa_Controle.frx":5663
         PICN            =   "Caixa_Controle.frx":567F
         PICH            =   "Caixa_Controle.frx":79D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdOKData 
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   540
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "OK"
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
         MICON           =   "Caixa_Controle.frx":9D25
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label12 
         Caption         =   "Data do Caixa"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.TextBox txtEntrada 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      TabIndex        =   12
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4515
      Left            =   60
      TabIndex        =   6
      Top             =   720
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   7964
      _Version        =   393216
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   60
      ScaleHeight     =   585
      ScaleWidth      =   12885
      TabIndex        =   4
      Top             =   60
      Width           =   12915
      Begin VB.Label lblCaixaRotulo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   12480
         TabIndex        =   111
         Top             =   180
         Width           =   270
      End
      Begin VB.Image Image1 
         Height          =   465
         Left            =   120
         Picture         =   "Caixa_Controle.frx":9D41
         Stretch         =   -1  'True
         Top             =   30
         Width           =   660
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHAMENTO DE CAIXA"
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
         Left            =   960
         TabIndex        =   5
         Top             =   150
         Width           =   3750
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Controle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   8640
      Width           =   8655
      Begin ChamaleonBtn.chameleonButton cmdFecharCaixa 
         Height          =   675
         Left            =   1740
         TabIndex        =   105
         Top             =   300
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Fechar Caixa"
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
         MICON           =   "Caixa_Controle.frx":10482
         PICN            =   "Caixa_Controle.frx":1049E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdMostrar 
         Height          =   675
         Left            =   60
         TabIndex        =   0
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1191
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
         MICON           =   "Caixa_Controle.frx":10905
         PICN            =   "Caixa_Controle.frx":10921
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAbrirCaixa 
         Height          =   675
         Left            =   1740
         TabIndex        =   1
         Top             =   300
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Abrir Caixa"
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
         MICON           =   "Caixa_Controle.frx":111FB
         PICN            =   "Caixa_Controle.frx":11217
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
         Height          =   675
         Left            =   5100
         TabIndex        =   2
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Imprimir Detralhado"
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
         MICON           =   "Caixa_Controle.frx":1167E
         PICN            =   "Caixa_Controle.frx":1169A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdTroco 
         Height          =   675
         Left            =   3420
         TabIndex        =   7
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Troco"
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
         MICON           =   "Caixa_Controle.frx":119B4
         PICN            =   "Caixa_Controle.frx":119D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImprimirResumido 
         Height          =   675
         Left            =   6780
         TabIndex        =   104
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Imprimir Resumido"
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
         MICON           =   "Caixa_Controle.frx":11B03
         PICN            =   "Caixa_Controle.frx":11B1F
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
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1920
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   28
      Top             =   9705
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10610
            Text            =   "Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "Caixa"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "Cód. Caixa"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Object.ToolTipText     =   "Situação"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "Data"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "18:19"
            Object.ToolTipText     =   "Hora"
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
   Begin ChamaleonBtn.chameleonButton chameleonButton1 
      Height          =   675
      Left            =   4080
      TabIndex        =   102
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1191
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
      MICON           =   "Caixa_Controle.frx":11E39
      PICN            =   "Caixa_Controle.frx":11E55
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdContarMoedas 
      Height          =   675
      Left            =   4080
      TabIndex        =   103
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1191
      BTYPE           =   3
      TX              =   "&Contar"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "Caixa_Controle.frx":1272F
      PICN            =   "Caixa_Controle.frx":1274B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdTrocarCaixa 
      Height          =   315
      Left            =   1380
      TabIndex        =   106
      Top             =   5280
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Trocar de Caixa"
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
      MICON           =   "Caixa_Controle.frx":14885
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdDetalhar 
      Height          =   315
      Left            =   60
      TabIndex        =   109
      Top             =   5280
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Detalhar"
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
      MICON           =   "Caixa_Controle.frx":148A1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCaixaPrincipal 
      Height          =   315
      Left            =   2700
      TabIndex        =   110
      Top             =   5280
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Caixa Principal"
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
      MICON           =   "Caixa_Controle.frx":148BD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblCodCaixaStatus 
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
      Left            =   6480
      TabIndex        =   112
      Top             =   5940
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblCodCaixaAtual 
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
      Left            =   6480
      TabIndex        =   108
      Top             =   5700
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblCaixaAtual 
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
      Left            =   6480
      TabIndex        =   107
      Top             =   5460
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "Caixa_Controle_semOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private printSQL As String
Private moCombo As cComboHelper

Dim varCodUsuario As Integer    'relatorio de impressao do caixa
Dim varNomeUsuario As String    'relatorio de impressao do caixa
Dim varDataFecha As Date        'relatorio de impressao do caixa
Dim varHoraFecha As String      'relatorio de impressao do caixa
'Dim var_Maquina As String       'colocar o nome da maquina na barra de status
Dim var_Caixa As String         'colocar o nome da maquina na barra de status
'Dim var_Setor As String         'mostrar o setor
Dim vStatusCaixaAtual As String

Dim varBotaoCaixa As Boolean
Dim varCodCaixa As Long
Dim vTipoCaixa As Integer

Dim vSaldoFisicoImpressão As Currency  'criei somente para impressão de caixa em branco
Dim vSaldoGeralImpressão As Currency  'criei somente para impressão de caixa em branco
Dim oCfg As ConfigItem

Dim vOSAtiva As Boolean
Dim vAluguelAtiva As Boolean
Dim i As Integer

Dim sSQL As String
Dim r As ADODB.Recordset
Private Sub CompararCaixa()
If lblCaixaAtual.Caption = var_Caixa Then
    cmdTrocarCaixa.Enabled = False
    cmdFecharCaixa.Enabled = True
    'cmdTroco.Enabled = True
Else
    cmdTrocarCaixa.Enabled = True
    cmdFecharCaixa.Enabled = False
    cmdTroco.Enabled = False
End If
End Sub

Private Sub EsconderTotais()
'faturamento
lblPrazo.Visible = False
txtQuantVendaPrazo.Visible = False
txtTotalVendaPrazo.Visible = False

lblAluguelFat.Visible = False
txtQuantAluguelPrazo.Visible = False
txtTotalAluguelPrazo.Visible = False

lblOSFat.Visible = False
txtQuantOSPrazo.Visible = False
txtTotalOSPrazo.Visible = False

'Saldo Fisico
lblSaldoFisico.Visible = True
txtSaldoFisico.Visible = True

lblVendas.Visible = False
txtQuantDinheiro.Visible = False
txtTotalDinheiro.Visible = False

lblOS.Visible = False
txtQuantDinheiroOS.Visible = False
txtTotalDinheiroOS.Visible = False

lblAluguel.Visible = False
txtQuantDinheiroAluguel.Visible = False
txtTotalDinheiroAluguel.Visible = False

lblParcelas.Visible = False
txtQuantDinheiroParcelas.Visible = False
txtTotalDinheiroParcelas.Visible = False

lblHaveres.Visible = False
txtQuantDinheiroHaveres.Visible = False
txtTotalDinheiroHaveres.Visible = False

lblSuprimentos.Visible = False
txtQuantDinheiroSuprimento.Visible = False
txtTotalDinheiroSuprimento.Visible = False

lblCheque.Visible = False
txtQuantCheque.Visible = False
txtTotalCheque.Visible = False

lblSaidas.Visible = False
txtQuantSaida.Visible = False
txtSaida.Visible = False

'saldo Geral
lblSaldoGeral.Visible = True
txtSaldo.Visible = True

lblCartao.Visible = False
txtQuantCartao.Visible = False
txtTotalCartao.Visible = False

lblOutros.Visible = False
txtQuantAvulso.Visible = False
txtTotalAvulso.Visible = False

lblSaldoFisico.Top = 900
txtSaldoFisico.Top = 900
cmdAbrirSaldoFisico.Top = 900

lblSaldoGeral.Top = 1200
txtSaldo.Top = 1200
cmdAbrirSaldoGeral.Top = 1200

lblFaturamento.Top = 1500
txtFaturamento.Top = 1500
cmdAbriFaturamento.Top = 1500

'escolher o que exibir no faturamento
If vAluguelAtiva = False And vOSAtiva = False Then
    lblFatServicos.Visible = False
    txtF7.Visible = False
    txtFATTotalServicos.Visible = False
    lblFatAluguel.Visible = False
    txtF8.Visible = False
    txtFATTotalAluguel.Visible = False
    lblSaldo.Top = 1980
    txtFATTotalSaldo.Top = 1980
    img10.Visible = False
    img11.Visible = False
    lblAluguelFat.Visible = False
    txtQuantAluguelPrazo.Visible = False
    txtTotalAluguelPrazo.Visible = False
    lblOSFat.Visible = False
    txtQuantOSPrazo.Visible = False
    txtTotalOSPrazo.Visible = False
ElseIf vAluguelAtiva = True And vOSAtiva = False Then
    lblFatServicos.Visible = False
    txtF7.Visible = False
    txtFATTotalServicos.Visible = False
    lblFatAluguel.Visible = True
    txtF8.Visible = True
    txtFATTotalAluguel.Visible = True
    lblFatAluguel.Top = 1980
    txtF8.Top = 1980
    txtFATTotalAluguel.Top = 1980
    lblSaldo.Top = 2280
    txtFATTotalSaldo.Top = 2280
    img11.Top = 1980
    img10.Visible = False
    'lblAluguelFat.Visible = True
    'txtQuantAluguelPrazo.Visible = True
    'txtTotalAluguelPrazo.Visible = True
    lblOSFat.Visible = False
    txtQuantOSPrazo.Visible = False
    txtTotalOSPrazo.Visible = False
ElseIf vAluguelAtiva = False And vOSAtiva = True Then
    lblFatServicos.Visible = True
    txtF7.Visible = True
    txtFATTotalServicos.Visible = True
    lblFatAluguel.Visible = False
    txtF8.Visible = False
    txtFATTotalAluguel.Visible = False
    lblFatServicos.Top = 1980
    txtF7.Top = 1980
    txtFATTotalServicos.Top = 1980
    lblSaldo.Top = 2280
    txtFATTotalSaldo.Top = 2280
    img10.Top = 1980
    img11.Visible = False
    lblAluguelFat.Visible = False
    txtQuantAluguelPrazo.Visible = False
    txtTotalAluguelPrazo.Visible = False
    'lblOSFat.Visible = True
    'txtQuantOSPrazo.Visible = True
    'txtTotalOSPrazo.Visible = True
ElseIf vAluguelAtiva = True And vOSAtiva = True Then
    lblFatServicos.Visible = True
    txtF7.Visible = True
    txtFATTotalServicos.Visible = True
    lblFatAluguel.Visible = True
    txtF8.Visible = True
    txtFATTotalAluguel.Visible = True
    lblFatServicos.Top = 1980
    txtF7.Top = 1980
    txtFATTotalServicos.Top = 1980
    lblFatAluguel.Top = 2280
    txtF8.Top = 2280
    txtFATTotalAluguel.Top = 2280
    lblSaldo.Top = 2580
    txtFATTotalSaldo.Top = 2580
    img10.Top = 1980
    img11.Top = 2280
    'lblAluguelFat.Visible = True
    'txtQuantAluguelPrazo.Visible = True
    'txtTotalAluguelPrazo.Visible = True
    'lblOSFat.Visible = True
    'txtQuantOSPrazo.Visible = True
    'txtTotalOSPrazo.Visible = True
End If
End Sub

Private Sub Mostrar_Retiradas()
sSQL = "SELECT ISNULL(SUM(VALOR), 0) AS vSomaRetiradas " & _
       "FROM caixa_retirada " & _
       "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") AND  caixa = '" & StatusBar1.Panels(2).Text & "' "
Set r = dbData.OpenRecordset(sSQL)

txtTotalRetiradas.Text = Format(r("vSomaRetiradas"), ocMONEY)


If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Mostrar_Servico()
Dim r_Prazo As ADODB.Recordset

sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaTotais " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & StatusBar1.Panels(3).Text & ") AND  pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'OFICINA') and pedidos.cancelado = 0 AND (parcelas.STATUS = 0)"
       'Debug.Print sSQL
Set r_Prazo = dbData.OpenRecordset(sSQL)

txtTotalOSPrazo.Text = Format(r_Prazo("varSomaTotais"), ocMONEY)
txtFATTotalServicos.Text = Format(r_Prazo("varSomaTotais"), ocMONEY)

''sSQL = "SELECT parcelas.VALOR_FINAL " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & StatusBar1.Panels(3).Text & ") AND  pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (parcelas.STATUS = 0)"
        'Debug.Print sSQL
        
sSQL = "SELECT pedidos.cod_pedido " & _
       "FROM pedidos INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
       "WHERE (pedidos.codcaixa = " & StatusBar1.Panels(3).Text & ") AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'OFICINA') AND pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' and pedidos.cancelado = 0"
'Debug.Print sSQL
Set r_Prazo = dbData.OpenRecordset(sSQL)

txtQuantOSPrazo.Text = Format(r_Prazo.RecordCount, "000")
txtF7.Text = Format(r_Prazo.RecordCount, "000")

If r_Prazo.State <> 0 Then r_Prazo.Close
Set r_Prazo = Nothing
End Sub
Private Sub Mostrar_Aluguel()
Dim r_Prazo As ADODB.Recordset

sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaTotais " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & StatusBar1.Panels(3).Text & ") AND  pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'ALUGUEL') and pedidos.cancelado = 0 AND (parcelas.STATUS = 0)"
       'Debug.Print sSQL
Set r_Prazo = dbData.OpenRecordset(sSQL)

txtTotalAluguelPrazo.Text = Format(r_Prazo("varSomaTotais"), ocMONEY)
txtFATTotalAluguel.Text = Format(r_Prazo("varSomaTotais"), ocMONEY)

sSQL = "SELECT pedidos.cod_pedido " & _
       "FROM pedidos INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
       "WHERE (pedidos.codcaixa = " & StatusBar1.Panels(3).Text & ") AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'ALUGUEL') AND pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' and pedidos.cancelado = 0"
       'Debug.Print sSQL
Set r_Prazo = dbData.OpenRecordset(sSQL)

txtQuantAluguelPrazo.Text = Format(r_Prazo.RecordCount, "000")
txtF8.Text = Format(r_Prazo.RecordCount, "000")

If r_Prazo.State <> 0 Then r_Prazo.Close
Set r_Prazo = Nothing
End Sub

Private Sub Mostrar_APrazo()
Dim r_Prazo As ADODB.Recordset

sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaTotais " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & StatusBar1.Panels(3).Text & ") AND  pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'VENDA') and pedidos.cancelado = 0 AND (parcelas.STATUS = 0)"
       'Debug.Print sSQL
Set r_Prazo = dbData.OpenRecordset(sSQL)

txtTotalVendaPrazo.Text = Format(r_Prazo("varSomaTotais"), ocMONEY)
txtFATTotalPrazo.Text = Format(r_Prazo("varSomaTotais"), ocMONEY)

''sSQL = "SELECT parcelas.VALOR_FINAL " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.codcaixa = " & StatusBar1.Panels(3).Text & ") AND  pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (parcelas.STATUS = 0)"
        'Debug.Print sSQL
        
sSQL = "SELECT pedidos.cod_pedido " & _
       "FROM pedidos INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
       "WHERE (pedidos.codcaixa = " & StatusBar1.Panels(3).Text & ") AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'VENDA') AND pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' and pedidos.cancelado = 0"
'Debug.Print sSQL
Set r_Prazo = dbData.OpenRecordset(sSQL)

txtQuantVendaPrazo.Text = Format(r_Prazo.RecordCount, "000")
txtF5.Text = Format(r_Prazo.RecordCount, "000")

If r_Prazo.State <> 0 Then r_Prazo.Close
Set r_Prazo = Nothing
End Sub
Private Sub Mostrar_Saldo()
Dim var_Troco As Currency
Dim var_Venda As Currency
Dim var_Parcela As Currency
Dim var_Haver As Currency
Dim var_Cheque As Currency
Dim var_Transf As Currency
Dim var_Suprimento As Currency
Dim var_Cartao As Currency
Dim var_Saida As Currency
Dim var_VendasPrazo As Currency
Dim var_AluguelPrazo As Currency
Dim var_SaldoFisico As Currency
Dim var_SaldoGeral As Currency
Dim var_Faturamento As Currency
Dim SaldoSemSaida As Currency
Dim var_OS As Currency
Dim var_ALUGUEL As Currency

'Inicializa as variáveis
var_Troco = 0
var_Venda = 0
var_Cheque = 0
var_Transf = 0
var_Cartao = 0
var_Suprimento = 0
var_Saida = 0
var_VendasPrazo = 0
var_OS = 0
var_ALUGUEL = 0
var_AluguelPrazo = 0

If chkTroco.Value = Checked Then
    If txtTotalTroco.Text <> "" Then var_Troco = txtTotalTroco.Text
Else
    var_Troco = 0
End If

If txtTotalDinheiro.Text <> "" Then var_Venda = txtTotalDinheiro.Text
If txtTotalDinheiroParcelas.Text <> "" Then var_Parcela = txtTotalDinheiroParcelas.Text
If txtTotalDinheiroHaveres.Text <> "" Then var_Haver = txtTotalDinheiroHaveres.Text
If txtTotalDinheiroSuprimento.Text <> "" Then var_Suprimento = txtTotalDinheiroSuprimento.Text
If txtTotalDinheiroOS.Text <> "" Then var_OS = txtTotalDinheiroOS.Text
If txtTotalDinheiroAluguel.Text <> "" Then var_ALUGUEL = txtTotalDinheiroAluguel.Text
If txtTotalCheque.Text <> "" Then var_Cheque = txtTotalCheque.Text
If txtSaida.Text <> "" Then var_Saida = txtSaida.Text
If txtTotalVendaPrazo <> "" Then var_VendasPrazo = txtTotalVendaPrazo.Text
If txtTotalAluguelPrazo <> "" Then var_AluguelPrazo = txtTotalAluguelPrazo.Text

'var_AluguelPrazo
'txtTotalAluguelPrazo

If txtTotalCartao.Text <> "" Then var_Cartao = txtTotalCartao.Text
If txtTotalAvulso.Text <> "" Then var_Transf = txtTotalAvulso.Text

var_SaldoFisico = var_Troco + var_Venda + var_Parcela + var_Haver + var_Cheque + var_Suprimento + var_ALUGUEL + var_OS
vSaldoFisicoImpressão = var_SaldoFisico  'criei somente para impressão de caixa em branco
var_SaldoFisico = var_SaldoFisico - var_Saida
txtSaldoFisico.Text = Format(var_SaldoFisico, ocMONEY)

var_SaldoGeral = var_SaldoFisico + var_Cartao + var_Transf
txtSaldo.Text = Format(var_SaldoGeral, ocMONEY)
vSaldoGeralImpressão = var_SaldoGeral  'criei somente para impressão de caixa em branco

SaldoSemSaida = var_SaldoGeral + var_Saida
txtEntrada.Text = Format(SaldoSemSaida, ocMONEY)

var_Faturamento = var_SaldoGeral + var_VendasPrazo + var_AluguelPrazo
txtFaturamento.Text = Format(var_Faturamento, ocMONEY)
End Sub



Private Sub SomaFlexCheque()
   On Error GoTo errorhandeler
   Dim soma As Currency
Dim QUANT As Integer
Dim i As Integer

soma = 0
QUANT = 0
   With Grid
      For i = 1 To .rows - 1
         If .TextMatrix(i, 5) = "CHEQUE" And IsNumeric(.TextMatrix(i, 6)) Then
            soma = soma + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
         End If
      Next
   End With
   
   txtTotalCheque.Text = Format(soma, ocMONEY)
   txtQuantCheque.Text = Format(QUANT, "000")
   
errorhandeler:
End Sub

Private Sub SomaFlexOutros()
On Error GoTo errorhandeler
Dim soma As Currency
Dim QUANT As Integer
Dim i As Integer

soma = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 5) = "DEPOSITO" Or .TextMatrix(i, 5) = "TRANSFERENCIA" Or .TextMatrix(i, 5) = "BOLETO" Or .TextMatrix(i, 5) = "FINANCEIRA" Or .TextMatrix(i, 5) = "PIX" And IsNumeric(.TextMatrix(i, 6)) Then
         soma = soma + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
      End If
   Next
End With

txtTotalAvulso.Text = Format(soma, ocMONEY)
txtQuantAvulso.Text = Format(QUANT, "000")
   
errorhandeler:
End Sub



Private Sub SomaFlexSaida()
   On Error GoTo errorhandeler
   Dim soma As Currency
Dim QUANT As Integer
Dim i As Integer

soma = 0
QUANT = 0
   With Grid
      For i = 1 To .rows - 1
         If .TextMatrix(i, 3) = "SANGRIA" And IsNumeric(.TextMatrix(i, 7)) Then
            soma = soma + CCur(.TextMatrix(i, 7))
         QUANT = QUANT + 1
         End If
      Next
   End With
   
   txtSaida.Text = Format(soma, "#,##0.00")
   txtQuantSaida.Text = Format(QUANT, "000")
   
errorhandeler:
End Sub

Private Sub SomaFlexCartao()
On Error GoTo errorhandeler
Dim soma As Currency
Dim QUANT As Integer
Dim i As Integer

soma = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If Left(.TextMatrix(i, 5), 6) = "CARTAO" And IsNumeric(.TextMatrix(i, 6)) Then
         soma = soma + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
      End If
   Next
End With

txtTotalCartao.Text = Format(soma, "#,##0.00")
txtQuantCartao.Text = Format(QUANT, "000")
   
errorhandeler:
End Sub

Private Sub SomaFaturamento()
On Error GoTo errorhandeler
Dim somaVendas As Currency
Dim somaParcelas As Currency
Dim somaHaveres As Currency
Dim somaSuprimentos As Currency
Dim somaSaidas As Currency
Dim somaPrazo As Currency
Dim Saldo As Currency
Dim somaOS As Currency
Dim somaAluguel As Currency
Dim QUANT As Integer
Dim i As Integer

'txtF7 = Format(CInt(txtQuantDinheiroOS) + CInt(txtQuantDinheiroAluguel), "000")
'txtFATTotalServicos = Format(CCur(txtTotalDinheiroOS) + CCur(txtTotalDinheiroAluguel), "#,##0.00")

'VENDAS
somaVendas = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 3) = "VENDA" And .TextMatrix(i, 10) <> "À Prazo" And IsNumeric(.TextMatrix(i, 6)) Then
         somaVendas = somaVendas + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
      End If
   Next
End With

txtFATTotalVendas.Text = Format(somaVendas, "#,##0.00")
txtF1.Text = Format(QUANT, "000")

'PARCELAS
somaParcelas = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 3) = "PARCELA" Or .TextMatrix(i, 3) = "ALUGUEL" Or .TextMatrix(i, 3) = "OS" Then
        If IsNumeric(.TextMatrix(i, 6)) Then
            somaParcelas = somaParcelas + CCur(.TextMatrix(i, 6))
            QUANT = QUANT + 1
        End If
      End If
   Next
End With

txtFATTotalParcelas.Text = Format(somaParcelas, "#,##0.00")
txtF2.Text = Format(QUANT, "000")

'HAVERES
somaHaveres = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 3) = "HAVER" And IsNumeric(.TextMatrix(i, 6)) Then
         somaHaveres = somaHaveres + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
      End If
   Next
End With

txtFATTotalHaveres.Text = Format(somaHaveres, "#,##0.00")
txtF3.Text = Format(QUANT, "000")

'SUPRIMENTOS
somaSuprimentos = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 3) = "SUPRIMENTO" And IsNumeric(.TextMatrix(i, 6)) Then
         somaSuprimentos = somaSuprimentos + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
      End If
   Next
End With

txtFATTotalSuprimentos.Text = Format(somaSuprimentos, "#,##0.00")
txtF4.Text = Format(QUANT, "000")

'==============================

'somaOS = 0
'QUANT = 0
'With Grid
'   For i = 1 To .Rows - 1
'      If .TextMatrix(i, 3) = "OS" And IsNumeric(.TextMatrix(i, 6)) Then
'         somaOS = somaOS + CCur(.TextMatrix(i, 6))
'         QUANT = QUANT + 1
'      End If
'   Next
'End With

'txtFATTotalSuprimentos.Text = Format(somaOS, "#,##0.00")
'txtF4.Text = Format(QUANT, "000")

'somaSuprimentos = 0
'QUANT = 0
'With Grid
'   For i = 1 To .Rows - 1
'      If .TextMatrix(i, 3) = "SUPRIMENTO" And IsNumeric(.TextMatrix(i, 6)) Then
'         somaSuprimentos = somaSuprimentos + CCur(.TextMatrix(i, 6))
'         QUANT = QUANT + 1
'      End If
'   Next
'End With
'
'txtFATTotalSuprimentos.Text = Format(somaSuprimentos, "#,##0.00")
'txtF4.Text = Format(QUANT, "000")


'======================
somaSaidas = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 3) = "SANGRIA" And IsNumeric(.TextMatrix(i, 7)) Then
         somaSaidas = somaSaidas + CCur(.TextMatrix(i, 7))
         QUANT = QUANT + 1
      End If
   Next
End With

txtFATTotalSaidas.Text = Format(somaSaidas, "#,##0.00")
txtF6.Text = Format(QUANT, "000")

somaPrazo = txtFATTotalPrazo.Text
somaOS = txtFATTotalServicos.Text
somaAluguel = txtFATTotalAluguel.Text

Saldo = somaVendas + somaParcelas + somaHaveres + somaSuprimentos + somaPrazo + somaOS + somaAluguel
Saldo = Saldo - somaSaidas
txtFATTotalSaldo.Text = Format(Saldo, "#,##0.00")

errorhandeler:
End Sub

Private Sub SomaFlexDinheiro()
On Error GoTo errorhandeler
Dim soma As Currency
Dim QUANT As Integer
Dim i As Integer

soma = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 5) = "DINHEIRO" And .TextMatrix(i, 3) = "VENDA" And IsNumeric(.TextMatrix(i, 6)) Then
         soma = soma + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
      End If
   Next
End With

txtTotalDinheiro.Text = Format(soma, "#,##0.00")
txtQuantDinheiro.Text = Format(QUANT, "000")

soma = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 5) = "DINHEIRO" And .TextMatrix(i, 3) = "PARCELA" And IsNumeric(.TextMatrix(i, 6)) Then
         soma = soma + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
      End If
   Next
End With

txtTotalDinheiroParcelas.Text = Format(soma, "#,##0.00")
txtQuantDinheiroParcelas.Text = Format(QUANT, "000")


soma = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 5) = "DINHEIRO" And .TextMatrix(i, 3) = "HAVER" And IsNumeric(.TextMatrix(i, 6)) Then
         soma = soma + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
      End If
   Next
End With

txtTotalDinheiroHaveres.Text = Format(soma, "#,##0.00")
txtQuantDinheiroHaveres.Text = Format(QUANT, "000")

soma = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 5) = "DINHEIRO" And .TextMatrix(i, 3) = "SUPRIMENTO" And IsNumeric(.TextMatrix(i, 6)) Then
         soma = soma + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
      End If
   Next
End With

txtTotalDinheiroSuprimento.Text = Format(soma, "#,##0.00")
txtQuantDinheiroSuprimento.Text = Format(QUANT, "000")

'==========================
soma = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 5) = "DINHEIRO" And .TextMatrix(i, 3) = "OS" And IsNumeric(.TextMatrix(i, 6)) Then
         soma = soma + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
      End If
   Next
End With

txtTotalDinheiroOS.Text = Format(soma, "#,##0.00")
txtQuantDinheiroOS.Text = Format(QUANT, "000")

'====================
soma = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 5) = "DINHEIRO" And .TextMatrix(i, 3) = "ALUGUEL" And IsNumeric(.TextMatrix(i, 6)) Then
         soma = soma + CCur(.TextMatrix(i, 6))
         QUANT = QUANT + 1
      End If
   Next
End With

txtTotalDinheiroAluguel.Text = Format(soma, "#,##0.00")
txtQuantDinheiroAluguel.Text = Format(QUANT, "000")
   
errorhandeler:
End Sub

Private Sub Limpa_Tudo()
   'txtQuantEnt.Text = ""
   'txtQuantCyb.Text = ""
   'txtQuantSaida.Text = ""
   'txtTotMens.Text = ""
   'txtTotCyb.Text = ""
   'txtTotSaida.Text = ""
   txtTotalDinheiro.Text = ""
   txtSaida.Text = ""
   'txtTotal.Text = ""
   'txtQuantGraf.Text = ""
   'txtTotGraf.Text = ""
   'txtTotDiv.Text = ""
   'txtQuantSai.Text = ""
End Sub

Private Sub Mostrar_Troco()
Dim TROCO As Currency

TROCO = 0
sSQL = "SELECT * FROM caixa_troco WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then TROCO = r("valor")
If r.State <> 0 Then r.Close
Set r = Nothing

txtTotalTroco.Text = Format(TROCO, ocMONEY)
End Sub

Private Sub VerificarCaixa()
If varFluxoCaixa = False Then
    sSQL = "SELECT *, CASE status WHEN 0 THEN 'ABERTO' ELSE 'FECHADO' END AS varStatus " & _
           "FROM caixa_dia " & _
           "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and caixa_dia.status = 0;"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.EOF Then
        varCodCaixa = ValidateNull(r("codcaixa"))
        cmdImprimir.Enabled = True
        cmdImprimirResumido.Enabled = True
        cmdAbrirCaixa.Visible = False
        cmdFecharCaixa.Visible = True
        cmdTroco.Enabled = True
        StatusBar1.Panels(3).Text = Format(ValidateNull(r("codcaixa")), "00000")
        StatusBar1.Panels(4).Text = r("VARSTATUS")
        lblCodCaixaAtual.Caption = Format(ValidateNull(r("codcaixa")), "00000")
        lblCodCaixaStatus.Caption = r("VARSTATUS")
    Else
        varCodCaixa = 0
        cmdTroco.Enabled = False
        cmdImprimir.Enabled = False
        cmdImprimirResumido.Enabled = False
        cmdAbrirCaixa.Visible = True
        cmdFecharCaixa.Visible = False
        cmdTroco.Enabled = False
        StatusBar1.Panels(3).Text = Format(0, "00000")
        StatusBar1.Panels(4).Text = "FECHADO"
    End If
    
    'mskData.Text = Format(Grid.TextMatrix(Grid.Row, 4), "dd/mm/yy")
    cmdMostrar_Click
End If
End Sub

Private Sub cboMaquina_GotFocus()
cboMaquina.Clear
cboMaquina.AddItem "CAIXA01"
cboMaquina.AddItem "CAIXA02"
cboMaquina.AddItem "CAIXA03"
cboMaquina.AddItem "TODOS"
moCombo.AttachTo cboMaquina
End Sub

Private Sub cboSetor_GotFocus()
cboSetor.Clear
cboSetor.AddItem "BALCAO"
cboSetor.AddItem "OFICINA"
cboSetor.AddItem "TODOS"
moCombo.AttachTo cboSetor
End Sub

Private Sub chameleonButton1_Click()
Dim SETOR_CAIXA As String
'Dim var_Setor As String
Dim varTipoCartao2 As String

If Not IsDate(mskData) Then Exit Sub

If varCodCaixa = 0 Then
    sSQL = "SELECT SUM(parcelas.valor_final) as vValorVendasTotal, 'VENDAS' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE 1=0"

    Set r = dbData.OpenRecordset(sSQL)
Else
    Dim Maquina_Parcela As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Parcela = "AND (parcelas.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Parcela = "AND (parcelas.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Haver As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Haver = "AND (parcelas_haver.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Haver = "AND (parcelas_haver.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Suprimento As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Suprimento = "AND (caixa_entrada.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Suprimento = "AND (caixa_entrada.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Sangria As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Sangria = "AND (caixa_saida.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Sangria = "AND (caixa_saida.caixa <> 'CAIXA') "
    End If
    
    SETOR_CAIXA = "AND (pedidos.tipo_pedido = 'VENDA') "

    
    'VENDAS
    sSQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'VENDAS' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de vendas - Dinheiro
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'VENDAS EM DINHEIRO' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'DINHEIRO' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de vendas - Pix
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'VENDAS EM PIX' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'PIX' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de vendas - Transferencia
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'VENDAS EM TRANSFERÊNCIA' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'TRANSFERENCIA' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de vendas - Deposito
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'VENDAS EM DEPOSITO' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'DEPOSITO' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de vendas - Financeira
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'VENDAS EM FINANCEIRA' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'FINANCEIRA' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de vendas - Cartão
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'VENDAS EM CARTÃO' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'CARTAO' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de vendas - Cheque
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'VENDAS EM CHEQUE' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'CHEQUE' " & Maquina_Parcela & _
           "UNION ALL "

    'PARCELAS
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'PARCELAS' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'PARCELA' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de Parcelas - Dinheiro
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'PARCELAS EM DINHEIRO' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'PARCELA' and FORMA_PGTO = 'DINHEIRO' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de Parcelas - Pix
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'PARCELAS EM PIX' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'PARCELA' and FORMA_PGTO = 'PIX' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de Parcelas - Transferencia
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'PARCELAS EM TRANSFERÊNCIA' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'PARCELA' and FORMA_PGTO = 'TRANSFERENCIA' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de Parcelas - Deposito
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'PARCELAS EM DEPOSITO' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'PARCELA' and FORMA_PGTO = 'DEPOSITO' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de Parcelas - Financeira
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'PARCELAS EM FINANCEIRA' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'PARCELA' and FORMA_PGTO = 'FINANCEIRA' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de Parcelas - Cartão
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'PARCELAS EM CARTÃO' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'PARCELA' and FORMA_PGTO = 'CARTAO' " & Maquina_Parcela & _
           "UNION ALL "
    'Detalhamento de Parcelas - Cheque
    sSQL = sSQL & "SELECT ISNULL(SUM(parcelas.valor_final),0) as vTotal, 'PARCELAS EM CHEQUE' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'PARCELA' and FORMA_PGTO = 'CHEQUE' " & Maquina_Parcela & _
           "UNION ALL "

    'HAVERES
    sSQL = sSQL & "SELECT ISNULL(SUM(VALOR_HAVER),0) as vTotal, 'HAVERES' as vTipoResultado FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' " & Maquina_Haver & _
           "UNION ALL "
    'Detalhamento de Haveres - Dinheiro
    sSQL = sSQL & "SELECT ISNULL(SUM(VALOR_HAVER),0) as vTotal, 'HAVERES EM DINHEIRO' as vTipoResultado FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'DINHEIRO' " & Maquina_Haver & _
           "UNION ALL "
    'Detalhamento de Haveres - Pix
    sSQL = sSQL & "SELECT ISNULL(SUM(VALOR_HAVER),0) as vTotal, 'HAVERES EM PIX' as vTipoResultado FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'PIX' " & Maquina_Haver & _
           "UNION ALL "
    'Detalhamento de Haveres - Transferencia
    sSQL = sSQL & "SELECT ISNULL(SUM(VALOR_HAVER),0) as vTotal, 'HAVERES EM TRANSFERÊNCIA' as vTipoResultado FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'TRANSFERENCIA' " & Maquina_Haver & _
           "UNION ALL "
    'Detalhamento de Haveres - Deposito
    sSQL = sSQL & "SELECT ISNULL(SUM(VALOR_HAVER),0) as vTotal, 'HAVERES EM DEPOSITO' as vTipoResultado FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'DEPOSITO' " & Maquina_Haver & _
           "UNION ALL "
    'Detalhamento de Haveres - Financeira
    sSQL = sSQL & "SELECT ISNULL(SUM(VALOR_HAVER),0) as vTotal, 'HAVERES EM FINANCEIRA' as vTipoResultado FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'FINANCEIRA' " & Maquina_Haver & _
           "UNION ALL "
    'Detalhamento de Haveres - Cartão
    sSQL = sSQL & "SELECT ISNULL(SUM(VALOR_HAVER),0) as vTotal, 'HAVERES EM CARTÃO' as vTipoResultado FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'CARTAO' " & Maquina_Haver & _
           "UNION ALL "
    'Detalhamento de Haveres - Cheque
    sSQL = sSQL & "SELECT ISNULL(SUM(VALOR_HAVER),0) as vTotal, 'HAVERES EM CHEQUE' as vTipoResultado FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'CHEQUE' " & Maquina_Haver & _
           "UNION ALL "

    'SUPRIMENTO
    sSQL = sSQL & "SELECT ISNULL(SUM(VALOR),0) as vTotal, 'SUPRIMENTOS' as vTipoResultado FROM caixa_entrada " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") " & Maquina_Suprimento & _
           "UNION ALL "

    'SANGRIA
    sSQL = sSQL & "SELECT ISNULL(SUM(VALOR),0) as vTotal, 'SANGRIAS' as vTipoResultado FROM caixa_saida " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") " & Maquina_Sangria
  
    'Debug.Print sSQL

    Set r = dbData.OpenRecordset(sSQL)
End If

'Mostrar_APrazo
'Mostrar_Retiradas

FormatarGridResumido r
  
If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL

End Sub

Private Sub chkTroco_Click()
Mostrar_Saldo
End Sub

Private Sub cmdAbriFaturamento_Click()
'faturamento
If cmdAbriFaturamento.Caption = "+" Then
    cmdAbriFaturamento.Caption = "-"
    cmdAbrirSaldoFisico.Caption = "+"
    cmdAbrirSaldoGeral.Caption = "+"
Else
    cmdAbriFaturamento.Caption = "+"
End If

If cmdAbriFaturamento.Caption = "-" Then
    lblPrazo.Visible = True
    txtQuantVendaPrazo.Visible = True
    txtTotalVendaPrazo.Visible = True
    
    If vAluguelAtiva = False And vOSAtiva = False Then
        lblAluguelFat.Visible = False
        txtQuantAluguelPrazo.Visible = False
        txtTotalAluguelPrazo.Visible = False
        lblOSFat.Visible = False
        txtQuantOSPrazo.Visible = False
        txtTotalOSPrazo.Visible = False
    ElseIf vAluguelAtiva = True And vOSAtiva = False Then
        lblAluguelFat.Visible = True
        txtQuantAluguelPrazo.Visible = True
        txtTotalAluguelPrazo.Visible = True
        lblOSFat.Visible = False
        txtQuantOSPrazo.Visible = False
        txtTotalOSPrazo.Visible = False
    ElseIf vAluguelAtiva = False And vOSAtiva = True Then
        lblAluguelFat.Visible = False
        txtQuantAluguelPrazo.Visible = False
        txtTotalAluguelPrazo.Visible = False
        lblOSFat.Visible = True
        txtQuantOSPrazo.Visible = True
        txtTotalOSPrazo.Visible = True
    ElseIf vAluguelAtiva = True And vOSAtiva = True Then
        lblAluguelFat.Visible = True
        txtQuantAluguelPrazo.Visible = True
        txtTotalAluguelPrazo.Visible = True
        lblOSFat.Visible = True
        txtQuantOSPrazo.Visible = True
        txtTotalOSPrazo.Visible = True
    End If
    
    lblSaldoFisico.Top = 900
    txtSaldoFisico.Top = 900
    cmdAbrirSaldoFisico.Top = 900
    
    lblSaldoGeral.Top = 1200
    txtSaldo.Top = 1200
    cmdAbrirSaldoGeral.Top = 1200
    
    lblFaturamento.Top = 1500
    txtFaturamento.Top = 1500
    cmdAbriFaturamento.Top = 1500
    
    lblPrazo.Top = 1800
    txtQuantVendaPrazo.Top = 1800
    txtTotalVendaPrazo.Top = 1800
    
    If vAluguelAtiva = False And vOSAtiva = False Then

    ElseIf vAluguelAtiva = True And vOSAtiva = False Then
        lblAluguelFat.Top = 2100
        txtQuantAluguelPrazo.Top = 2100
        txtTotalAluguelPrazo.Top = 2100
    ElseIf vAluguelAtiva = False And vOSAtiva = True Then
        lblOSFat.Top = 2100
        txtQuantOSPrazo.Top = 2100
        txtTotalOSPrazo.Top = 2100
    ElseIf vAluguelAtiva = True And vOSAtiva = True Then
        lblAluguelFat.Top = 2100
        txtQuantAluguelPrazo.Top = 2100
        txtTotalAluguelPrazo.Top = 2100
        lblOSFat.Top = 2400
        txtQuantOSPrazo.Top = 2400
        txtTotalOSPrazo.Top = 2400
    End If
    
Else
    lblPrazo.Visible = False
    txtQuantVendaPrazo.Visible = False
    txtTotalVendaPrazo.Visible = False
    
    lblAluguelFat.Visible = False
    txtQuantAluguelPrazo.Visible = False
    txtTotalAluguelPrazo.Visible = False

    lblOSFat.Visible = False
    txtQuantOSPrazo.Visible = False
    txtTotalOSPrazo.Visible = False
    
    lblSaldoFisico.Top = 900
    txtSaldoFisico.Top = 900
    cmdAbrirSaldoFisico.Top = 900
    
    lblSaldoGeral.Top = 1200
    txtSaldo.Top = 1200
    cmdAbrirSaldoGeral.Top = 1200
    
    lblFaturamento.Top = 1500
    txtFaturamento.Top = 1500
    cmdAbriFaturamento.Top = 1500
    
End If


'Saldo Fisico
lblSaldoFisico.Visible = True
txtSaldoFisico.Visible = True

lblVendas.Visible = False
txtQuantDinheiro.Visible = False
txtTotalDinheiro.Visible = False

lblOS.Visible = False
txtQuantDinheiroOS.Visible = False
txtTotalDinheiroOS.Visible = False

lblAluguel.Visible = False
txtQuantDinheiroAluguel.Visible = False
txtTotalDinheiroAluguel.Visible = False

lblParcelas.Visible = False
txtQuantDinheiroParcelas.Visible = False
txtTotalDinheiroParcelas.Visible = False

lblHaveres.Visible = False
txtQuantDinheiroHaveres.Visible = False
txtTotalDinheiroHaveres.Visible = False

lblSuprimentos.Visible = False
txtQuantDinheiroSuprimento.Visible = False
txtTotalDinheiroSuprimento.Visible = False

lblCheque.Visible = False
txtQuantCheque.Visible = False
txtTotalCheque.Visible = False

lblSaidas.Visible = False
txtQuantSaida.Visible = False
txtSaida.Visible = False

'saldo Geral
lblSaldoGeral.Visible = True
txtSaldo.Visible = True

lblCartao.Visible = False
txtQuantCartao.Visible = False
txtTotalCartao.Visible = False

lblOutros.Visible = False
txtQuantAvulso.Visible = False
txtTotalAvulso.Visible = False

img7.Visible = False
img8.Visible = False



End Sub
Private Sub cmdAbrirCaixa_Click()
'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
'cmdMostrar_Click

'If cmdAbrirCaixa.Visible = True Then
'   frmSenha.Visible = True
'   txtSenha.SetFocus
'End If
varBotaoCaixa = True
Caixa_Fechamento.Show 1
End Sub

Private Sub cmdAbrirSaldoFisico_Click()
If cmdAbrirSaldoFisico.Caption = "+" Then
    cmdAbrirSaldoFisico.Caption = "-"
    cmdAbriFaturamento.Caption = "+"
    cmdAbrirSaldoGeral.Caption = "+"
Else
    cmdAbrirSaldoFisico.Caption = "+"
End If


If cmdAbrirSaldoFisico.Caption = "-" Then
    lblVendas.Visible = True
    txtQuantDinheiro.Visible = True
    txtTotalDinheiro.Visible = True
    
    lblParcelas.Visible = True
    txtQuantDinheiroParcelas.Visible = True
    txtTotalDinheiroParcelas.Visible = True
    
    lblHaveres.Visible = True
    txtQuantDinheiroHaveres.Visible = True
    txtTotalDinheiroHaveres.Visible = True
    
    lblSuprimentos.Visible = True
    txtQuantDinheiroSuprimento.Visible = True
    txtTotalDinheiroSuprimento.Visible = True
    
    lblCheque.Visible = True
    txtQuantCheque.Visible = True
    txtTotalCheque.Visible = True
    
    lblSaidas.Visible = True
    txtQuantSaida.Visible = True
    txtSaida.Visible = True

    If vOSAtiva = True And vAluguelAtiva = False Then
    
        lblOS.Visible = True
        txtQuantDinheiroOS.Visible = True
        txtTotalDinheiroOS.Visible = True

        lblAluguel.Visible = False
        txtQuantDinheiroAluguel.Visible = False
        txtTotalDinheiroAluguel.Visible = False

        lblSaldoFisico.Top = 900
        txtSaldoFisico.Top = 900
        cmdAbrirSaldoFisico.Top = 900
        
        lblVendas.Top = 1200
        txtQuantDinheiro.Top = 1200
        txtTotalDinheiro.Top = 1200
        
        lblOS.Top = 1500
        txtQuantDinheiroOS.Top = 1500
        txtTotalDinheiroOS.Top = 1500
        
        lblParcelas.Top = 1800
        txtQuantDinheiroParcelas.Top = 1800
        txtTotalDinheiroParcelas.Top = 1800
        
        lblHaveres.Top = 2100
        txtQuantDinheiroHaveres.Top = 2100
        txtTotalDinheiroHaveres.Top = 2100
        
        lblSuprimentos.Top = 2400
        txtQuantDinheiroSuprimento.Top = 2400
        txtTotalDinheiroSuprimento.Top = 2400
        
        lblCheque.Top = 2700
        txtQuantCheque.Top = 2700
        txtTotalCheque.Top = 2700
        
        lblSaidas.Top = 3000
        txtQuantSaida.Top = 3000
        txtSaida.Top = 3000
        
        lblSaldoGeral.Top = 3300
        txtSaldo.Top = 3300
        cmdAbrirSaldoGeral.Top = 3300
        
        lblFaturamento.Top = 3600
        txtFaturamento.Top = 3600
        cmdAbriFaturamento.Top = 3600
        
    ElseIf vOSAtiva = False And vAluguelAtiva = True Then
    
        lblOS.Visible = False
        txtQuantDinheiroOS.Visible = False
        txtTotalDinheiroOS.Visible = False

        lblAluguel.Visible = True
        txtQuantDinheiroAluguel.Visible = True
        txtTotalDinheiroAluguel.Visible = True

        lblSaldoFisico.Top = 900
        txtSaldoFisico.Top = 900
        cmdAbrirSaldoFisico.Top = 900
        
        lblVendas.Top = 1200
        txtQuantDinheiro.Top = 1200
        txtTotalDinheiro.Top = 1200
        
        lblAluguel.Top = 1500
        txtQuantDinheiroAluguel.Top = 1500
        txtTotalDinheiroAluguel.Top = 1500
        
        lblParcelas.Top = 1800
        txtQuantDinheiroParcelas.Top = 1800
        txtTotalDinheiroParcelas.Top = 1800
        
        lblHaveres.Top = 2100
        txtQuantDinheiroHaveres.Top = 2100
        txtTotalDinheiroHaveres.Top = 2100
        
        lblSuprimentos.Top = 2400
        txtQuantDinheiroSuprimento.Top = 2400
        txtTotalDinheiroSuprimento.Top = 2400
        
        lblCheque.Top = 2700
        txtQuantCheque.Top = 2700
        txtTotalCheque.Top = 2700
        
        lblSaidas.Top = 3000
        txtQuantSaida.Top = 3000
        txtSaida.Top = 3000
        
        lblSaldoGeral.Top = 3300
        txtSaldo.Top = 3300
        cmdAbrirSaldoGeral.Top = 3300
        
        lblFaturamento.Top = 3600
        txtFaturamento.Top = 3600
        cmdAbriFaturamento.Top = 3600
        
    ElseIf vOSAtiva = False And vAluguelAtiva = False Then
    
        lblOS.Visible = False
        txtQuantDinheiroOS.Visible = False
        txtTotalDinheiroOS.Visible = False

        lblAluguel.Visible = False
        txtQuantDinheiroAluguel.Visible = False
        txtTotalDinheiroAluguel.Visible = False

        lblSaldoFisico.Top = 900
        txtSaldoFisico.Top = 900
        cmdAbrirSaldoFisico.Top = 900
        
        lblVendas.Top = 1200
        txtQuantDinheiro.Top = 1200
        txtTotalDinheiro.Top = 1200
        
        lblParcelas.Top = 1500
        txtQuantDinheiroParcelas.Top = 1500
        txtTotalDinheiroParcelas.Top = 1500
        
        lblHaveres.Top = 1800
        txtQuantDinheiroHaveres.Top = 1800
        txtTotalDinheiroHaveres.Top = 1800
        
        lblSuprimentos.Top = 2100
        txtQuantDinheiroSuprimento.Top = 2100
        txtTotalDinheiroSuprimento.Top = 2100
        
        lblCheque.Top = 2400
        txtQuantCheque.Top = 2400
        txtTotalCheque.Top = 2400
        
        lblSaidas.Top = 2700
        txtQuantSaida.Top = 2700
        txtSaida.Top = 2700
        
        lblSaldoGeral.Top = 3000
        txtSaldo.Top = 3000
        cmdAbrirSaldoGeral.Top = 3000
        
        lblFaturamento.Top = 3300
        txtFaturamento.Top = 3300
        cmdAbriFaturamento.Top = 3300
        
    ElseIf vOSAtiva = True And vAluguelAtiva = True Then

        lblOS.Visible = True
        txtQuantDinheiroOS.Visible = True
        txtTotalDinheiroOS.Visible = True

        lblAluguel.Visible = True
        txtQuantDinheiroAluguel.Visible = True
        txtTotalDinheiroAluguel.Visible = True

        lblSaldoFisico.Top = 900
        txtSaldoFisico.Top = 900
        cmdAbrirSaldoFisico.Top = 900
        
        lblVendas.Top = 1200
        txtQuantDinheiro.Top = 1200
        txtTotalDinheiro.Top = 1200
        
        lblOS.Top = 1500
        txtQuantDinheiroOS.Top = 1500
        txtTotalDinheiroOS.Top = 1500
        
        lblAluguel.Top = 1800
        txtQuantDinheiroAluguel.Top = 1800
        txtTotalDinheiroAluguel.Top = 1800
        
        lblParcelas.Top = 2100
        txtQuantDinheiroParcelas.Top = 2100
        txtTotalDinheiroParcelas.Top = 2100
        
        lblHaveres.Top = 2400
        txtQuantDinheiroHaveres.Top = 2400
        txtTotalDinheiroHaveres.Top = 2400
        
        lblSuprimentos.Top = 2700
        txtQuantDinheiroSuprimento.Top = 2700
        txtTotalDinheiroSuprimento.Top = 2700
        
        lblCheque.Top = 3000
        txtQuantCheque.Top = 3000
        txtTotalCheque.Top = 3000
        
        lblSaidas.Top = 3300
        txtQuantSaida.Top = 3300
        txtSaida.Top = 3300
        
        lblSaldoGeral.Top = 3600
        txtSaldo.Top = 3600
        cmdAbrirSaldoGeral.Top = 3600
        
        lblFaturamento.Top = 3900
        txtFaturamento.Top = 3900
        cmdAbriFaturamento.Top = 3900

    End If
        
   

Else
    lblVendas.Visible = False
    txtQuantDinheiro.Visible = False
    txtTotalDinheiro.Visible = False
    
    lblOS.Visible = False
    txtQuantDinheiroOS.Visible = False
    txtTotalDinheiroOS.Visible = False
    
    lblAluguel.Visible = False
    txtQuantDinheiroAluguel.Visible = False
    txtTotalDinheiroAluguel.Visible = False
    
    lblParcelas.Visible = False
    txtQuantDinheiroParcelas.Visible = False
    txtTotalDinheiroParcelas.Visible = False
    
    lblHaveres.Visible = False
    txtQuantDinheiroHaveres.Visible = False
    txtTotalDinheiroHaveres.Visible = False
    
    lblSuprimentos.Visible = False
    txtQuantDinheiroSuprimento.Visible = False
    txtTotalDinheiroSuprimento.Visible = False
    
    lblCheque.Visible = False
    txtQuantCheque.Visible = False
    txtTotalCheque.Visible = False
    
    lblSaidas.Visible = False
    txtQuantSaida.Visible = False
    txtSaida.Visible = False

    lblSaldoFisico.Top = 900
    txtSaldoFisico.Top = 900
    cmdAbrirSaldoFisico.Top = 900
    
    lblSaldoGeral.Top = 1200
    txtSaldo.Top = 1200
    cmdAbrirSaldoGeral.Top = 1200
    
    lblFaturamento.Top = 1500
    txtFaturamento.Top = 1500
    cmdAbriFaturamento.Top = 1500
End If

'faturamento
lblPrazo.Visible = False
txtQuantVendaPrazo.Visible = False
txtTotalVendaPrazo.Visible = False

lblAluguelFat.Visible = False
txtQuantAluguelPrazo.Visible = False
txtTotalAluguelPrazo.Visible = False

lblOSFat.Visible = False
txtQuantOSPrazo.Visible = False
txtTotalOSPrazo.Visible = False

'Saldo Fisico
'lblSaldoFisico.Visible = True
'txtSaldoFisico.Visible = True

'lblVendas.Visible = True
'txtQuantDinheiro.Visible = True
'txtTotalDinheiro.Visible = True

'lblOS.Visible = True
'txtQuantDinheiroOS.Visible = True
'txtTotalDinheiroOS.Visible = True

'lblAluguel.Visible = True
'txtQuantDinheiroAluguel.Visible = True
'txtTotalDinheiroAluguel.Visible = True

'lblParcelas.Visible = True
'txtQuantDinheiroParcelas.Visible = True
'txtTotalDinheiroParcelas.Visible = True

'lblHaveres.Visible = True
'txtQuantDinheiroHaveres.Visible = True
'txtTotalDinheiroHaveres.Visible = True

'lblSuprimentos.Visible = True
'txtQuantDinheiroSuprimento.Visible = True
'txtTotalDinheiroSuprimento.Visible = True

'lblCheque.Visible = True
'txtQuantCheque.Visible = True
'txtTotalCheque.Visible = True

'lblSaidas.Visible = True
'txtQuantSaida.Visible = True
'txtSaida.Visible = True

'saldo Geral
lblSaldoGeral.Visible = True
txtSaldo.Visible = True

lblFaturamento.Visible = True
txtFaturamento.Visible = True
cmdAbriFaturamento.Visible = True

lblCartao.Visible = False
txtQuantCartao.Visible = False
txtTotalCartao.Visible = False

lblOutros.Visible = False
txtQuantAvulso.Visible = False
txtTotalAvulso.Visible = False

img7.Visible = False
img8.Visible = False

'POSIÇÃO ===========================


End Sub

Private Sub cmdAbrirSaldoGeral_Click()
If cmdAbrirSaldoGeral.Caption = "+" Then
    cmdAbrirSaldoGeral.Caption = "-"
    cmdAbriFaturamento.Caption = "+"
    cmdAbrirSaldoFisico.Caption = "+"
Else
    cmdAbrirSaldoGeral.Caption = "+"
End If

If cmdAbrirSaldoGeral.Caption = "-" Then
    lblSaldoGeral.Visible = True
    txtSaldo.Visible = True
    
    lblCartao.Visible = True
    txtQuantCartao.Visible = True
    txtTotalCartao.Visible = True
    
    lblOutros.Visible = True
    txtQuantAvulso.Visible = True
    txtTotalAvulso.Visible = True
    
    lblSaldoFisico.Top = 900
    txtSaldoFisico.Top = 900
    cmdAbrirSaldoFisico.Top = 900

    lblSaldoGeral.Top = 1200
    txtSaldo.Top = 1200
    cmdAbrirSaldoGeral.Top = 1200
    
    lblCartao.Top = 1500
    txtQuantCartao.Top = 1500
    txtTotalCartao.Top = 1500
    
    lblOutros.Top = 1800
    txtQuantAvulso.Top = 1800
    txtTotalAvulso.Top = 1800

    lblFaturamento.Top = 2100
    txtFaturamento.Top = 2100
    cmdAbriFaturamento.Top = 2100

Else
    lblSaldoGeral.Visible = True
    txtSaldo.Visible = True
    
    lblCartao.Visible = False
    txtQuantCartao.Visible = False
    txtTotalCartao.Visible = False
    
    lblOutros.Visible = False
    txtQuantAvulso.Visible = False
    txtTotalAvulso.Visible = False
    
    lblSaldoFisico.Top = 900
    txtSaldoFisico.Top = 900
    cmdAbrirSaldoFisico.Top = 900

    lblSaldoGeral.Top = 1200
    txtSaldo.Top = 1200
    cmdAbrirSaldoGeral.Top = 1200

    lblFaturamento.Top = 1500
    txtFaturamento.Top = 1500
    cmdAbriFaturamento.Top = 1500
End If

'faturamento
lblPrazo.Visible = False
txtQuantVendaPrazo.Visible = False
txtTotalVendaPrazo.Visible = False

lblAluguelFat.Visible = False
txtQuantAluguelPrazo.Visible = False
txtTotalAluguelPrazo.Visible = False
lblOSFat.Visible = False
txtQuantOSPrazo.Visible = False
txtTotalOSPrazo.Visible = False

'Saldo Fisico
lblSaldoFisico.Visible = True
txtSaldoFisico.Visible = True

lblVendas.Visible = False
txtQuantDinheiro.Visible = False
txtTotalDinheiro.Visible = False

lblOS.Visible = False
txtQuantDinheiroOS.Visible = False
txtTotalDinheiroOS.Visible = False

lblAluguel.Visible = False
txtQuantDinheiroAluguel.Visible = False
txtTotalDinheiroAluguel.Visible = False

lblParcelas.Visible = False
txtQuantDinheiroParcelas.Visible = False
txtTotalDinheiroParcelas.Visible = False

lblHaveres.Visible = False
txtQuantDinheiroHaveres.Visible = False
txtTotalDinheiroHaveres.Visible = False

lblSuprimentos.Visible = False
txtQuantDinheiroSuprimento.Visible = False
txtTotalDinheiroSuprimento.Visible = False

lblCheque.Visible = False
txtQuantCheque.Visible = False
txtTotalCheque.Visible = False

lblSaidas.Visible = False
txtQuantSaida.Visible = False
txtSaida.Visible = False

'saldo Geral
'lblSaldoGeral.Visible = True
'txtSaldo.Visible = True

'lblCartao.Visible = True
'txtQuantCartao.Visible = True
'txtTotalCartao.Visible = True

'lblOutros.Visible = True
'txtQuantAvulso.Visible = True
'txtTotalAvulso.Visible = True

'lblSaldoFisico.Top = 1200
'txtSaldoFisico.Top = 1200
'cmdAbrirSaldoFisico.Top = 1200

'lblCartao.Top = 1800
'txtQuantCartao.Top = 1800
'txtTotalCartao.Top = 1800

'lblOutros.Top = 2100
'txtQuantAvulso.Top = 2100
'txtTotalAvulso.Top = 2100

'lblSaldoGeral.Top = 1500
'txtSaldo.Top = 1500
'cmdAbrirSaldoGeral.Top = 1500



End Sub

Private Sub cmdCaixaPrincipal_Click()
If cmdCaixaPrincipal.Caption = "Caixa Principal" Then
    StatusBar1.Panels(2).Text = "CAIXA01"
    var_Caixa = "CAIXA01"
    lblCaixaRotulo.Caption = "CAIXA01"
    Call MostrarCodCaixa
    Call cmdMostrar_Click
    cmdCaixaPrincipal.Caption = "Voltar ao Caixa"
ElseIf cmdCaixaPrincipal.Caption = "Voltar ao Caixa" Then
    StatusBar1.Panels(2).Text = lblCaixaAtual.Caption
    var_Caixa = lblCaixaAtual.Caption
    lblCaixaRotulo.Caption = lblCaixaAtual.Caption
    Call MostrarCodCaixa
    Call cmdMostrar_Click
    cmdCaixaPrincipal.Caption = "Caixa Principal"
End If
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

Private Sub cmdDetalhar_Click()
If Not IsNumeric(Grid.TextMatrix(Grid.Row, 2)) = True Then Exit Sub
If Grid.TextMatrix(Grid.Row, 2) = "" Or Grid.TextMatrix(Grid.Row, 3) = "" Then Exit Sub

If Grid.TextMatrix(Grid.Row, 3) = "OS" Then
   Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 2), "OS"
Else
   Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 2), Grid.TextMatrix(Grid.Row, 3)
End If
Parcelas_Consulta_Produtos.Show 1
End Sub

Private Sub cmdFecharCaixa_Click()
'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
chkTroco.Value = Unchecked
'chkCartao.Value = Checked
'chkTransf.Value = Checked
'cmdMostrar_Click
Mostrar_Saldo
'frmSenha.Visible = True
'txtSenha.SetFocus
varBotaoCaixa = True
Load Caixa_Fechamento
Caixa_Fechamento.lblTitulo.Caption = "FECHAMENTO DO CAIXA"
Caixa_Fechamento.txtTroco.Enabled = False
Caixa_Fechamento.lblTroco.Enabled = False
Caixa_Fechamento.Show 1
End Sub

Public Sub cmdImprimir_Click()
'If StatusBar1.Panels(4).Text <> "FECHADO" Then MsgBox "Não é permitido imprimir um caixa aberto!", vbInformation, "Aviso do Sistema": Exit Sub

Dim r As ADODB.Recordset
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

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

'PEGAR OS DADOS DO FECHAMENTO
Dim sSQLusuario As String
Dim r_usuario As ADODB.Recordset

sSQLusuario = "SELECT DATA_ABERTURA, HORA_ABERTURA, COD_FUNC_ABERTURA, DATA_FECHAMENTO, HORA_FECHAMENTO, COD_FUNC_FECHAMENTO, (CASE WHEN status = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS VarStatus, " & _
        "(SELECT Usuario.Login FROM Usuario INNER JOIN caixa_dia ON Usuario.Codigo = caixa_dia.COD_FUNC_ABERTURA wHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & StatusBar1.Panels(3).Text & ")) AS Nome_Abertura, " & _
        "(SELECT Usuario_2.Login FROM Usuario AS Usuario_2 INNER JOIN caixa_dia AS caixa_dia_2 ON Usuario_2.Codigo = caixa_dia_2.COD_FUNC_FECHAMENTO WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & StatusBar1.Panels(3).Text & ")) AS Nome_Fechamento " & _
       "FROM caixa_dia AS caixa_dia_1 " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & StatusBar1.Panels(3).Text & ");"
Set r_usuario = dbData.OpenRecordset(sSQLusuario)

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

If vSaldoFisicoImpressão = "0" And vSaldoGeralImpressão = "0" Then    'fiz esse if para imprimir caixa sem saldo
    If r.State <> 0 Then r.Close
    Set r = Nothing
End If

If vAluguelAtiva = False And vOSAtiva = False Then
    Set REL_Caixa_Fech_Imp.ReportMain1.Recordset = r
    
    REL_Caixa_Fech_Imp.txtDHead.Caption = "FECHAMENTO DE CAIXA - ABERTURA: " & Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
    
    'REL_Caixa_Fech_Imp.rfTroco.Caption = Format(txtTotalTroco.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp.rfDinheiro.Caption = Format(txtTotalDinheiro.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfParcelas.Caption = Format(txtTotalDinheiroParcelas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfHaveres.Caption = Format(txtTotalDinheiroHaveres.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfSuprimentos.Caption = Format(txtTotalDinheiroSuprimento.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfCheque.Caption = Format(txtTotalCheque.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfSaida.Caption = Format(txtSaida.Text, "#,##0.00") & " "
    'REL_Caixa_Fech_Imp.rfAluguel2.Caption = Format(txtTotalDinheiroAluguel.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfSaldoFisico.Caption = Format(txtSaldoFisico.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp.rfCartao.Caption = Format(txtTotalCartao.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfOutros.Caption = Format(txtTotalAvulso.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfSaldoGeral.Caption = Format(txtSaldo.Text, "#,##0.00") & " "
    
    'REL_Caixa_Fech_Imp.rfAluguel.Caption = Format(txtTotalAluguelPrazo.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfPrazo.Caption = Format(txtTotalVendaPrazo.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfFaturamento.Caption = Format(txtFaturamento.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp.rfDinheiroQuant.Caption = Format(txtQuantDinheiro, "000") & " "
    'REL_Caixa_Fech_Imp.rfAluguelQuant2.Caption = Format(txtQuantDinheiroAluguel, "000") & " "
    REL_Caixa_Fech_Imp.rfParcelasQuant.Caption = Format(txtQuantDinheiroParcelas, "000") & " "
    REL_Caixa_Fech_Imp.rfHaveresQuant.Caption = Format(txtQuantDinheiroHaveres, "000") & " "
    REL_Caixa_Fech_Imp.rfSuprimentosQuant.Caption = Format(txtQuantDinheiroSuprimento, "000") & " "
    REL_Caixa_Fech_Imp.rfChequeQuant.Caption = Format(txtQuantCheque, "000") & " "
    REL_Caixa_Fech_Imp.rfSaidaQuant.Caption = Format(txtQuantSaida, "000") & " "
    REL_Caixa_Fech_Imp.rfCartaoQuant.Caption = Format(txtQuantCartao, "000") & " "
    REL_Caixa_Fech_Imp.rfOutrosQuant.Caption = Format(txtQuantAvulso, "000") & " "
    REL_Caixa_Fech_Imp.rfPrazoQuant.Caption = Format(txtQuantVendaPrazo, "000") & " "
    'REL_Caixa_Fech_Imp.rfAluguelQuant.Caption = Format(txtQuantDinheiroAluguel, "000") & " "
    
    
    REL_Caixa_Fech_Imp.rfT1.Caption = Format(txtFATTotalVendas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfT2.Caption = Format(txtFATTotalParcelas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfT3.Caption = Format(txtFATTotalHaveres.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfT4.Caption = Format(txtFATTotalSuprimentos.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfT5.Caption = Format(txtFATTotalPrazo.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfT6.Caption = Format(txtFATTotalSaidas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp.rfFTotal.Caption = Format(txtFATTotalSaldo.Text, "#,##0.00") & " "
    'REL_Caixa_Fech_Imp.rfT7.Caption = Format(txtFATTotalAluguel.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp.rfF1.Caption = Format(txtF1, "000") & " "
    REL_Caixa_Fech_Imp.rfF2.Caption = Format(txtF2, "000") & " "
    REL_Caixa_Fech_Imp.rfF3.Caption = Format(txtF3, "000") & " "
    REL_Caixa_Fech_Imp.rfF4.Caption = Format(txtF4, "000") & " "
    REL_Caixa_Fech_Imp.rfF5.Caption = Format(txtF5, "000") & " "
    REL_Caixa_Fech_Imp.rfF6.Caption = Format(txtF6, "000") & " "
    'REL_Caixa_Fech_Imp.rfF7.Caption = Format(txtF8, "000") & " "
    
    '===========================RODAPÉ
    If Not r_usuario.EOF Then
        REL_Caixa_Fech_Imp.rfCodUsuarioA.Caption = Format(r_usuario("COD_FUNC_ABERTURA"), "00")
        REL_Caixa_Fech_Imp.rfNomeUsuarioA.Caption = ValidateNull(r_usuario("Nome_Abertura"))
        REL_Caixa_Fech_Imp.rfDataA.Caption = Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
        REL_Caixa_Fech_Imp.rfHoraA.Caption = Format(ValidateNull(r_usuario("HORA_ABERTURA")), "hh:mm")
        
        REL_Caixa_Fech_Imp.rfNomeUsuarioF.Caption = ValidateNull(r_usuario("Nome_Fechamento"))
        If IsNull(r_usuario("DATA_FECHAMENTO")) Then
            REL_Caixa_Fech_Imp.rfDataF.Caption = ""
            REL_Caixa_Fech_Imp.rfCodUsuarioF.Caption = ""
            REL_Caixa_Fech_Imp.rfHoraF.Caption = ""
        Else
            REL_Caixa_Fech_Imp.rfCodUsuarioF.Caption = Format(ValidateNull(r_usuario("COD_FUNC_FECHAMENTO")), "00")
            REL_Caixa_Fech_Imp.rfDataF.Caption = Format(ValidateNull(r_usuario("DATA_FECHAMENTO")), "dd/mm/yyyy")
            REL_Caixa_Fech_Imp.rfHoraF.Caption = Format(ValidateNull(r_usuario("HORA_FECHAMENTO")), "hh:mm")
        End If
    
        REL_Caixa_Fech_Imp.rfSituacao.Caption = ValidateNull(r_usuario("VARSTATUS"))
    End If
    
    REL_Caixa_Fech_Imp.rfCaixa.Caption = StatusBar1.Panels(2).Text
    REL_Caixa_Fech_Imp.rfCodCaixa.Caption = Format(StatusBar1.Panels(3).Text, "0000")
    
    REL_Caixa_Fech_Imp.ReportMain1.NomeImpressora = var_ImpNormal
    REL_Caixa_Fech_Imp.ReportMain1.Ativar
    Unload REL_Caixa_Fech_Imp

ElseIf vAluguelAtiva = True And vOSAtiva = False Then
    Set REL_Caixa_Fech_Imp_Aluguel.ReportMain1.Recordset = r
    
    REL_Caixa_Fech_Imp_Aluguel.txtDHead.Caption = "FECHAMENTO DE CAIXA - ABERTURA: " & Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
    
    'REL_Caixa_Fech_Imp_Aluguel.rfTroco.Caption = Format(txtTotalTroco.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp_Aluguel.rfDinheiro.Caption = Format(txtTotalDinheiro.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfParcelas.Caption = Format(txtTotalDinheiroParcelas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfHaveres.Caption = Format(txtTotalDinheiroHaveres.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfSuprimentos.Caption = Format(txtTotalDinheiroSuprimento.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfCheque.Caption = Format(txtTotalCheque.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfSaida.Caption = Format(txtSaida.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfAluguel2.Caption = Format(txtTotalDinheiroAluguel.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfSaldoFisico.Caption = Format(txtSaldoFisico.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp_Aluguel.rfCartao.Caption = Format(txtTotalCartao.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfOutros.Caption = Format(txtTotalAvulso.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfSaldoGeral.Caption = Format(txtSaldo.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp_Aluguel.rfAluguel.Caption = Format(txtTotalAluguelPrazo.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfPrazo.Caption = Format(txtTotalVendaPrazo.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfFaturamento.Caption = Format(txtFaturamento.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp_Aluguel.rfDinheiroQuant.Caption = Format(txtQuantDinheiro, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfAluguelQuant2.Caption = Format(txtQuantDinheiroAluguel, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfParcelasQuant.Caption = Format(txtQuantDinheiroParcelas, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfHaveresQuant.Caption = Format(txtQuantDinheiroHaveres, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfSuprimentosQuant.Caption = Format(txtQuantDinheiroSuprimento, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfChequeQuant.Caption = Format(txtQuantCheque, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfSaidaQuant.Caption = Format(txtQuantSaida, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfCartaoQuant.Caption = Format(txtQuantCartao, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfOutrosQuant.Caption = Format(txtQuantAvulso, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfPrazoQuant.Caption = Format(txtQuantVendaPrazo, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfAluguelQuant.Caption = Format(txtF8, "000") & " "
    
    
    REL_Caixa_Fech_Imp_Aluguel.rfT1.Caption = Format(txtFATTotalVendas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfT2.Caption = Format(txtFATTotalParcelas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfT3.Caption = Format(txtFATTotalHaveres.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfT4.Caption = Format(txtFATTotalSuprimentos.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfT5.Caption = Format(txtFATTotalPrazo.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfT6.Caption = Format(txtFATTotalSaidas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfFTotal.Caption = Format(txtFATTotalSaldo.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfT7.Caption = Format(txtFATTotalAluguel.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp_Aluguel.rfF1.Caption = Format(txtF1, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfF2.Caption = Format(txtF2, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfF3.Caption = Format(txtF3, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfF4.Caption = Format(txtF4, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfF5.Caption = Format(txtF5, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfF6.Caption = Format(txtF6, "000") & " "
    REL_Caixa_Fech_Imp_Aluguel.rfF7.Caption = Format(txtF8, "000") & " "
    
    '===========================RODAPÉ
    If Not r_usuario.EOF Then
        REL_Caixa_Fech_Imp_Aluguel.rfCodUsuarioA.Caption = Format(r_usuario("COD_FUNC_ABERTURA"), "00")
        REL_Caixa_Fech_Imp_Aluguel.rfNomeUsuarioA.Caption = ValidateNull(r_usuario("Nome_Abertura"))
        REL_Caixa_Fech_Imp_Aluguel.rfDataA.Caption = Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
        REL_Caixa_Fech_Imp_Aluguel.rfHoraA.Caption = Format(ValidateNull(r_usuario("HORA_ABERTURA")), "hh:mm")
        
        REL_Caixa_Fech_Imp_Aluguel.rfNomeUsuarioF.Caption = ValidateNull(r_usuario("Nome_Fechamento"))
        If IsNull(r_usuario("DATA_FECHAMENTO")) Then
            REL_Caixa_Fech_Imp_Aluguel.rfDataF.Caption = ""
            REL_Caixa_Fech_Imp_Aluguel.rfCodUsuarioF.Caption = ""
            REL_Caixa_Fech_Imp_Aluguel.rfHoraF.Caption = ""
        Else
            REL_Caixa_Fech_Imp_Aluguel.rfCodUsuarioF.Caption = Format(ValidateNull(r_usuario("COD_FUNC_FECHAMENTO")), "00")
            REL_Caixa_Fech_Imp_Aluguel.rfDataF.Caption = Format(ValidateNull(r_usuario("DATA_FECHAMENTO")), "dd/mm/yyyy")
            REL_Caixa_Fech_Imp_Aluguel.rfHoraF.Caption = Format(ValidateNull(r_usuario("HORA_FECHAMENTO")), "hh:mm")
        End If
    
        REL_Caixa_Fech_Imp_Aluguel.rfSituacao.Caption = ValidateNull(r_usuario("VARSTATUS"))
    End If
    
    REL_Caixa_Fech_Imp_Aluguel.rfCaixa.Caption = StatusBar1.Panels(2).Text
    REL_Caixa_Fech_Imp_Aluguel.rfCodCaixa.Caption = Format(StatusBar1.Panels(3).Text, "0000")
    
    REL_Caixa_Fech_Imp_Aluguel.ReportMain1.NomeImpressora = var_ImpNormal
    REL_Caixa_Fech_Imp_Aluguel.ReportMain1.Ativar
    Unload REL_Caixa_Fech_Imp_Aluguel
ElseIf vAluguelAtiva = False And vOSAtiva = True Then
    Set REL_Caixa_Fech_Imp_OS.ReportMain1.Recordset = r
    
    REL_Caixa_Fech_Imp_OS.txtDHead.Caption = "FECHAMENTO DE CAIXA - ABERTURA: " & Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
    
    'REL_Caixa_Fech_Imp_OS.rfTroco.Caption = Format(txtTotalTroco.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp_OS.rfDinheiro.Caption = Format(txtTotalDinheiro.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfParcelas.Caption = Format(txtTotalDinheiroParcelas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfHaveres.Caption = Format(txtTotalDinheiroHaveres.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfSuprimentos.Caption = Format(txtTotalDinheiroSuprimento.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfCheque.Caption = Format(txtTotalCheque.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfSaida.Caption = Format(txtSaida.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfOS2.Caption = Format(txtTotalDinheiroOS.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfSaldoFisico.Caption = Format(txtSaldoFisico.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp_OS.rfCartao.Caption = Format(txtTotalCartao.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfOutros.Caption = Format(txtTotalAvulso.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfSaldoGeral.Caption = Format(txtSaldo.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp_OS.rfOS.Caption = Format(txtTotalOSPrazo.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfPrazo.Caption = Format(txtTotalVendaPrazo.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfFaturamento.Caption = Format(txtFaturamento.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp_OS.rfDinheiroQuant.Caption = Format(txtQuantDinheiro, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfOSQuant2.Caption = Format(txtQuantDinheiroOS, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfParcelasQuant.Caption = Format(txtQuantDinheiroParcelas, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfHaveresQuant.Caption = Format(txtQuantDinheiroHaveres, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfSuprimentosQuant.Caption = Format(txtQuantDinheiroSuprimento, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfChequeQuant.Caption = Format(txtQuantCheque, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfSaidaQuant.Caption = Format(txtQuantSaida, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfCartaoQuant.Caption = Format(txtQuantCartao, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfOutrosQuant.Caption = Format(txtQuantAvulso, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfPrazoQuant.Caption = Format(txtQuantVendaPrazo, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfOSQuant.Caption = Format(txtF8, "000") & " "
    
    
    REL_Caixa_Fech_Imp_OS.rfT1.Caption = Format(txtFATTotalVendas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfT2.Caption = Format(txtFATTotalParcelas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfT3.Caption = Format(txtFATTotalHaveres.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfT4.Caption = Format(txtFATTotalSuprimentos.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfT5.Caption = Format(txtFATTotalPrazo.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfT6.Caption = Format(txtFATTotalSaidas.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfFTotal.Caption = Format(txtFATTotalSaldo.Text, "#,##0.00") & " "
    REL_Caixa_Fech_Imp_OS.rfT7.Caption = Format(txtFATTotalServicos.Text, "#,##0.00") & " "
    
    REL_Caixa_Fech_Imp_OS.rfF1.Caption = Format(txtF1, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfF2.Caption = Format(txtF2, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfF3.Caption = Format(txtF3, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfF4.Caption = Format(txtF4, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfF5.Caption = Format(txtF5, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfF6.Caption = Format(txtF6, "000") & " "
    REL_Caixa_Fech_Imp_OS.rfF7.Caption = Format(txtF8, "000") & " "
    
    '===========================RODAPÉ
    If Not r_usuario.EOF Then
        REL_Caixa_Fech_Imp_OS.rfCodUsuarioA.Caption = Format(r_usuario("COD_FUNC_ABERTURA"), "00")
        REL_Caixa_Fech_Imp_OS.rfNomeUsuarioA.Caption = ValidateNull(r_usuario("Nome_Abertura"))
        REL_Caixa_Fech_Imp_OS.rfDataA.Caption = Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
        REL_Caixa_Fech_Imp_OS.rfHoraA.Caption = Format(ValidateNull(r_usuario("HORA_ABERTURA")), "hh:mm")
        
        REL_Caixa_Fech_Imp_OS.rfNomeUsuarioF.Caption = ValidateNull(r_usuario("Nome_Fechamento"))
        If IsNull(r_usuario("DATA_FECHAMENTO")) Then
            REL_Caixa_Fech_Imp_OS.rfDataF.Caption = ""
            REL_Caixa_Fech_Imp_OS.rfCodUsuarioF.Caption = ""
            REL_Caixa_Fech_Imp_OS.rfHoraF.Caption = ""
        Else
            REL_Caixa_Fech_Imp_OS.rfCodUsuarioF.Caption = Format(ValidateNull(r_usuario("COD_FUNC_FECHAMENTO")), "00")
            REL_Caixa_Fech_Imp_OS.rfDataF.Caption = Format(ValidateNull(r_usuario("DATA_FECHAMENTO")), "dd/mm/yyyy")
            REL_Caixa_Fech_Imp_OS.rfHoraF.Caption = Format(ValidateNull(r_usuario("HORA_FECHAMENTO")), "hh:mm")
        End If
    
        REL_Caixa_Fech_Imp_OS.rfSituacao.Caption = ValidateNull(r_usuario("VARSTATUS"))
    End If
    
    REL_Caixa_Fech_Imp_OS.rfCaixa.Caption = StatusBar1.Panels(2).Text
    REL_Caixa_Fech_Imp_OS.rfCodCaixa.Caption = Format(StatusBar1.Panels(3).Text, "0000")
    
    REL_Caixa_Fech_Imp_OS.ReportMain1.NomeImpressora = var_ImpNormal
    REL_Caixa_Fech_Imp_OS.ReportMain1.Ativar
    Unload REL_Caixa_Fech_Imp_OS
End If
Me.Show
End Sub

Private Sub cmdImprimirResumido_Click()
Dim SETOR_CAIXA As String
'Dim var_Setor As String
Dim varTipoCartao2 As String
Dim SQL As String

If Not IsDate(mskData) Then Exit Sub

'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

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

'    If vOSAtiva = True And vAluguelAtiva = False Then

If varCodCaixa = 0 Then
    SQL = "SELECT SUM(parcelas.valor_final) as vValorVendasTotal, 'VENDAS' as vTipoResultado FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE 1=0"
    Set r = dbData.OpenRecordset(SQL)
Else
    Dim Maquina_Parcela As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Parcela = "AND (parcelas.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Parcela = "AND (parcelas.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Venda As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Venda = "AND (caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Venda = "AND (caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Haver As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Haver = "AND (parcelas_haver.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Haver = "AND (parcelas_haver.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Suprimento As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Suprimento = "AND (caixa_entrada.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Suprimento = "AND (caixa_entrada.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Sangria As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Sangria = "AND (caixa_saida.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Sangria = "AND (caixa_saida.caixa <> 'CAIXA') "
    End If
    
    SETOR_CAIXA = "AND (pedidos.tipo_pedido = 'VENDA') "
    
    'TROCO
    Dim vVlrTroco As Currency
    '"SELECT * FROM caixa_troco WHERE (caixa_troco.codcaixa = " & StatusBar1.Panels(3).Text & ") AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
    SQL = "SELECT * FROM caixa_troco WHERE (caixa_troco.codcaixa = " & StatusBar1.Panels(3).Text & ") AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrTroco = r("VALOR") Else vVlrTroco = 0
    
    'VENDAS
    Dim vVlrVendasTotal As Currency
    Dim vQtdeVendasTotal As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasTotal, count(codigo) as vQuantVendasTotal FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasTotal = r("vValorVendasTotal"): vQtdeVendasTotal = r("vQuantVendasTotal") Else vVlrVendasTotal = 0: vQtdeVendasTotal = 0

    'Detalhamento de vendas - Dinheiro
    Dim vVlrVendasDinheiro As Currency
    Dim vQtdeVendasDinheiro As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasDinheiro, count(codigo) as vQuantVendasDinheiro FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'DINHEIRO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasDinheiro = r("vValorVendasDinheiro"): vQtdeVendasDinheiro = r("vQuantVendasDinheiro") Else vVlrVendasDinheiro = 0: vQtdeVendasDinheiro = 0


    'Detalhamento de vendas - Pix
    Dim vVlrVendasPix As Currency
    Dim vQtdeVendasPix As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasPix, count(codigo) as vQuantVendasPix FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'PIX' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasPix = r("vValorVendasPix"): vQtdeVendasPix = r("vQuantVendasPix") Else vVlrVendasPix = 0: vQtdeVendasPix = 0

    'Detalhamento de vendas - Transferencia
    Dim vVlrVendasTransferencia As Currency
    Dim vQtdeVendasTransferencia As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasTransferencia, count(codigo) as vQuantVendasTransferencia FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'TRANSFERENCIA' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasTransferencia = r("vValorVendasTransferencia"): vQtdeVendasTransferencia = r("vQuantVendasTransferencia") Else vVlrVendasTransferencia = 0: vQtdeVendasTransferencia = 0

    'Detalhamento de vendas - Deposito
    Dim vVlrVendasDeposito As Currency
    Dim vQtdeVendasDeposito As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasDeposito, count(codigo) as vQuantVendasDeposito FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'DEPOSITO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasDeposito = r("vValorVendasDeposito"): vQtdeVendasDeposito = r("vQuantVendasDeposito") Else vVlrVendasDeposito = 0: vQtdeVendasDeposito = 0

    'Detalhamento de vendas - Financeira
    Dim vVlrVendasFinanceira As Currency
    Dim vQtdeVendasFinanceira As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasFinanceira, count(codigo) as vQuantVendasFinanceira FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'FINANCEIRA' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasFinanceira = r("vValorVendasFinanceira"): vQtdeVendasFinanceira = r("vQuantVendasFinanceira") Else vVlrVendasFinanceira = 0: vQtdeVendasFinanceira = 0

    'Detalhamento de vendas - Cartão
    Dim vVlrVendasCartao As Currency
    Dim vQtdeVendasCartao As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasCartao, count(codigo) as vQuantVendasCartao FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'CARTAO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasCartao = r("vValorVendasCartao"): vQtdeVendasCartao = r("vQuantVendasCartao") Else vVlrVendasCartao = 0: vQtdeVendasCartao = 0

    'Detalhamento de vendas - Cheque
    Dim vVlrVendasCheque As Currency
    Dim vQtdeVendasCheque As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorVendasCheque, count(codigo) as vQuantVendasCheque FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and parcelas.tipo = 'VENDA' and FORMA_PGTO = 'CHEQUE' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrVendasCheque = r("vValorVendasCheque"): vQtdeVendasCheque = r("vQuantVendasCheque") Else vVlrVendasCheque = 0: vQtdeVendasCheque = 0


    'PARCELAS
    Dim vVlrParcelasTotal As Currency
    Dim vQtdeParcelasTotal As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaTotal, count(codigo) as vQuantParcelaTotal FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasTotal = r("vValorParcelaTotal"): vQtdeParcelasTotal = r("vQuantParcelaTotal") Else vVlrParcelasTotal = 0: vQtdeParcelasTotal = 0

    'Detalhamento de Parcelas - Dinheiro
    Dim vVlrParcelasDinheiro As Currency
    Dim vQtdeParcelasDinheiro As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaDinheiro, count(codigo) as vQuantParcelaDinheiro FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'DINHEIRO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    'Debug.Print SQL
    If Not r.EOF Then vVlrParcelasDinheiro = r("vValorParcelaDinheiro"): vQtdeParcelasDinheiro = r("vQuantParcelaDinheiro") Else vVlrParcelasDinheiro = 0: vQtdeParcelasDinheiro = 0

    'Detalhamento de Parcelas - Pix
    Dim vVlrParcelasPix As Currency
    Dim vQtdeParcelasPix As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaPix, count(codigo) as vQuantParcelaPix FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'PIX' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasPix = r("vValorParcelaPix"): vQtdeParcelasPix = r("vQuantParcelaPix") Else vVlrParcelasPix = 0: vQtdeVendasTotal = 0: vQtdeParcelasPix = 0

    'Detalhamento de Parcelas - Transferencia
    Dim vVlrParcelasTransferencia As Currency
    Dim vQtdeParcelasTransferencia As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaTransferencia, count(codigo) as vQuantParcelaTransferencia FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'TRANSFERENCIA' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasTransferencia = r("vValorParcelaTransferencia"): vQtdeParcelasTransferencia = r("vQuantParcelaTransferencia") Else vVlrParcelasTransferencia = 0: vQtdeParcelasTransferencia = 0

    'Detalhamento de Parcelas - Deposito
    Dim vVlrParcelasDeposito As Currency
    Dim vQtdeParcelasDeposito As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaDeposito, count(codigo) as vQuantParcelaDeposito FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'DEPOSITO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasDeposito = r("vValorParcelaDeposito"): vQtdeParcelasDeposito = r("vQuantParcelaDeposito") Else vVlrParcelasDeposito = 0: vQtdeParcelasDeposito = 0

    'Detalhamento de Parcelas - Financeira
    Dim vVlrParcelasFinanceira As Currency
    Dim vQtdeParcelasFinanceira As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaFinanceira, count(codigo) as vQuantParcelaFinanceira FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'FINANCEIRA' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasFinanceira = r("vValorParcelaFinanceira"): vQtdeParcelasFinanceira = r("vQuantParcelaFinanceira") Else vVlrParcelasFinanceira = 0: vQtdeParcelasFinanceira = 0

    'Detalhamento de Parcelas - Cartão
    Dim vVlrParcelasCartao As Currency
    Dim vQtdeParcelasCartao As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaCartao, count(codigo) as vQuantParcelaCartao FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'CARTAO' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasCartao = r("vValorParcelaCartao"): vQtdeParcelasCartao = r("vQuantParcelaCartao") Else vVlrParcelasCartao = 0: vQtdeParcelasCartao = 0

    'Detalhamento de Parcelas - Cheque
    Dim vVlrParcelasCheque As Currency
    Dim vQtdeParcelasCheque As Integer
    SQL = "SELECT ISNULL(SUM(parcelas.valor_final),0) as vValorParcelaCheque, count(codigo) as vQuantParcelaCheque FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") and (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) and FORMA_PGTO = 'CHEQUE' " & Maquina_Parcela
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrParcelasCheque = r("vValorParcelaCheque"): vQtdeParcelasCheque = r("vQuantParcelaCheque") Else vVlrParcelasCheque = 0: vQtdeParcelasCheque = 0


    'HAVERES
    Dim vVlrHaveresTotal As Currency
    Dim vQtdeHaveresTotal As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresTotal, count(codigo) as vQuantHaveresTotal FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresTotal = r("vValorHaveresTotal"): vQtdeHaveresTotal = r("vQuantHaveresTotal") Else vVlrHaveresTotal = 0: vQtdeHaveresTotal = 0
    
    'Detalhamento de Haveres - Dinheiro
    Dim vVlrHaveresDinheiro As Currency
    Dim vQtdeHaveresDinheiro As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresDinheiro, count(codigo) as vQuantHaveresDinheiro FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'DINHEIRO' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresDinheiro = r("vValorHaveresDinheiro"): vQtdeHaveresDinheiro = r("vQuantHaveresDinheiro") Else vVlrHaveresDinheiro = 0: vQtdeHaveresDinheiro = 0

    'Detalhamento de Haveres - Pix
    Dim vVlrHaveresPix As Currency
    Dim vQtdeHaveresPix As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresPix, count(codigo) as vQuantHaveresPix FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'PIX' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresPix = r("vValorHaveresPix"): vQtdeHaveresPix = r("vQuantHaveresPix") Else vVlrHaveresPix = 0: vQtdeHaveresPix = 0

    'Detalhamento de Haveres - Transferencia
    Dim vVlrHaveresTransferencia As Currency
    Dim vQtdeHaveresTransferencia As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresTransferencia, count(codigo) as vQuantHaveresTransferencia FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'TRANSFERENCIA' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresTransferencia = r("vValorHaveresTransferencia"): vQtdeHaveresTransferencia = r("vQuantHaveresTransferencia") Else vVlrHaveresTransferencia = 0: vQtdeHaveresTransferencia = 0

    'Detalhamento de Haveres - Deposito
    Dim vVlrHaveresDeposito As Currency
    Dim vQtdeHaveresDeposito As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresDeposito, count(codigo) as vQuantHaveresDeposito FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'DEPOSITO' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresDeposito = r("vValorHaveresDeposito"): vQtdeHaveresDeposito = r("vQuantHaveresDeposito") Else vVlrHaveresDeposito = 0: vQtdeHaveresDeposito = 0

    'Detalhamento de Haveres - Financeira
    Dim vVlrHaveresFinanceira As Currency
    Dim vQtdeHaveresFinanceira As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresFinanceira, count(codigo) as vQuantHaveresFinanceira FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'FINANCEIRA' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresFinanceira = r("vValorHaveresFinanceira"): vQtdeHaveresFinanceira = r("vQuantHaveresFinanceira") Else vVlrHaveresFinanceira = 0: vQtdeHaveresFinanceira = 0

    'Detalhamento de Haveres - Cartão
    Dim vVlrHaveresCartao As Currency
    Dim vQtdeHaveresCartao As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresCartao, count(codigo) as vQuantHaveresCartao FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'CARTAO' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresCartao = r("vValorHaveresCartao"): vQtdeHaveresCartao = r("vQuantHaveresCartao") Else vVlrHaveresCartao = 0: vQtdeHaveresCartao = 0

    'Detalhamento de Haveres - Cheque
    Dim vVlrHaveresCheque As Currency
    Dim vQtdeHaveresCheque As Integer
    SQL = "SELECT ISNULL(SUM(VALOR_HAVER),0) as vValorHaveresCheque, count(codigo) as vQuantHaveresCheque FROM parcelas_haver " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") and tipo = 'PARCELA' and FORMA_PGTO = 'CHEQUE' " & Maquina_Haver
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrHaveresCheque = r("vValorHaveresCheque"): vQtdeHaveresCheque = r("vQuantHaveresCheque") Else vVlrHaveresCheque = 0: vQtdeHaveresCheque = 0


    'SUPRIMENTO
    Dim vVlrSuprimento As Currency
    Dim vQtdeSuprimento As Integer
    SQL = "SELECT ISNULL(SUM(VALOR),0) as vValorSuprimentoTotal, count(codigo) as vQuantSuprimentoTotal FROM caixa_entrada " & _
           "WHERE (codcaixa = " & StatusBar1.Panels(3).Text & ") " & Maquina_Suprimento
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrSuprimento = r("vValorSuprimentoTotal"): vQtdeSuprimento = r("vQuantSuprimentoTotal") Else vVlrSuprimento = 0: vQtdeSuprimento = 0


    'SANGRIA
    Dim vVlrSangria As Currency
    Dim vQtdeSangria As Integer
    SQL = "SELECT ISNULL(SUM(VALOR),0) as vValorSangriaTotal, count(codigo) as vQuantSangriaTotal FROM caixa_saida " & _
           "WHERE FONTE = 'CAIXA ATUAL' AND (codcaixa = " & StatusBar1.Panels(3).Text & ") " & Maquina_Sangria
    Set r = dbData.OpenRecordset(SQL)
    If Not r.EOF Then vVlrSangria = r("vValorSangriaTotal"): vQtdeSangria = r("vQuantSangriaTotal") Else vVlrSangria = 0: vQtdeSangria = 0

    
    
    'VENDAS A PRAZO
    Dim vVlrVendasPrazoTotal As Currency
    Dim vQtdeVendasPrazoTotal As Integer
    sSQL = "SELECT ISNULL(SUM(TOTAL), 0) AS varSomaPrazoTotais, count(cod_pedido) as varQuantPrazoTotal " & _
           "FROM pedidos  " & _
           "WHERE TIPO_PAGAMENTO = 'À Prazo' and TIPO_PEDIDO= 'VENDA' and pedidos.cancelado = 0 AND (codcaixa = " & StatusBar1.Panels(3).Text & ") " & Maquina_Venda
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vQtdeVendasPrazoTotal = r("varQuantPrazoTotal") Else vQtdeVendasPrazoTotal = 0

    sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaPrazoTotais, count(pedidos.cod_pedido) as varQuantPrazoTotal " & _
       "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
       "WHERE (pedidos.tipo_pagamento = 'À Prazo') and (pedidos.TIPO_PEDIDO = 'VENDA') and pedidos.cancelado = 0 and (pedidos.codcaixa = " & StatusBar1.Panels(3).Text & ") AND  pedidos.caixa = '" & StatusBar1.Panels(2).Text & "'  AND (parcelas.STATUS = 0)"
    
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vVlrVendasPrazoTotal = r("varSomaPrazoTotais") Else vVlrVendasPrazoTotal = 0

    'VENDAS CANCELADAS
    Dim vVlrVendasCanceladoTotal As Currency
    Dim vQtdeVendasCanceladoTotal As Integer
    sSQL = "SELECT ISNULL(SUM(TOTAL), 0) AS varSomaCanceladoTotais, count(cod_pedido) as varQuantCanceladoTotal " & _
           "FROM pedidos  " & _
           "WHERE pedidos.cancelado = 1 AND (codcaixa = " & StatusBar1.Panels(3).Text & ") " & Maquina_Venda
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vVlrVendasCanceladoTotal = r("varSomaCanceladoTotais"): vQtdeVendasCanceladoTotal = r("varQuantCanceladoTotal") Else vVlrVendasCanceladoTotal = 0: vQtdeVendasCanceladoTotal = 0

   'ORÇAMENTO
    Dim vVlrVendasOrcamentoTotal As Currency
    Dim vQtdeVendasOrcamentoTotal As Integer
    sSQL = "SELECT ISNULL(SUM(TOTAL), 0) AS varSomaOrcamentoTotais, count(cod_pedido) as varQuantOrcamentoTotal " & _
           "FROM pedidos  " & _
           "WHERE pedidos.cancelado = 1 AND (codcaixa = " & StatusBar1.Panels(3).Text & ") " & Maquina_Venda
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vVlrVendasOrcamentoTotal = r("varSomaOrcamentoTotais"): vQtdeVendasOrcamentoTotal = r("varQuantOrcamentoTotal") Else vVlrVendasOrcamentoTotal = 0: vQtdeVendasOrcamentoTotal = 0

   'CONSIGNADO
    Dim vVlrVendasConsignadoTotal As Currency
    Dim vQtdeVendasConsignadoTotal As Integer
    sSQL = "SELECT ISNULL(SUM(TOTAL), 0) AS varSomaConsignadoTotais, count(cod_pedido) as varQuantConsignadoTotal " & _
           "FROM pedidos  " & _
           "WHERE TIPO_PEDIDO= 'CONSIGNADO' AND pedidos.cancelado = 0 AND (codcaixa = " & StatusBar1.Panels(3).Text & ") " & Maquina_Venda
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.EOF Then vVlrVendasConsignadoTotal = r("varSomaConsignadoTotais"): vQtdeVendasConsignadoTotal = r("varQuantConsignadoTotal") Else vVlrVendasConsignadoTotal = 0: vQtdeVendasConsignadoTotal = 0

    'ALUGUEL
    Dim vVlrAluguelTotal As Currency
    Dim vQtdeAluguelTotal As Integer
    sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaTotaisAluguel, count(pedidos.COD_PEDIDO) as varQuantAluguelTotal " & _
           "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
           "WHERE (pedidos.codcaixa = " & StatusBar1.Panels(3).Text & ") AND  pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'ALUGUEL') and pedidos.cancelado = 0 AND (parcelas.STATUS = 0)"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.EOF Then vVlrAluguelTotal = r("varSomaTotaisAluguel"): vQtdeAluguelTotal = r("varQuantAluguelTotal") Else vVlrAluguelTotal = 0: vQtdeAluguelTotal = 0

    'OS
    Dim vVlrOSTotal As Currency
    Dim vQtdeOSTotal As Integer
    sSQL = "SELECT ISNULL(SUM(parcelas.VALOR_FINAL), 0) AS varSomaTotaisOS, count(pedidos.COD_PEDIDO) as varQuantOSTotal " & _
           "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & _
           "WHERE (pedidos.codcaixa = " & StatusBar1.Panels(3).Text & ") AND  pedidos.caixa = '" & StatusBar1.Panels(2).Text & "' AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'OFICINA') and pedidos.cancelado = 0 AND (parcelas.STATUS = 0)"
    Set r = dbData.OpenRecordset(sSQL)
    Debug.Print sSQL
    
    If Not r.EOF Then vVlrOSTotal = r("varSomaTotaisOS"): vQtdeOSTotal = r("varQuantOSTotal") Else vVlrOSTotal = 0: vQtdeOSTotal = 0




    'Set r = dbData.OpenRecordset(SQL)
End If


'Mostrar_APrazo
'Mostrar_Retiradas

'FormatarGridResumido r
  
'If r.State <> 0 Then r.Close
'Set r = Nothing

'mostrar todas as saídas na folha
SQL = "SELECT HORA as vSHora, SUBDESCRICAO + '/' + DESCRICAO as vSDescricao, COD_FUNCIONARIO as vSFunc, VALOR as vSValor FROM caixa_saida " & _
       "WHERE FONTE = 'CAIXA ATUAL' AND (codcaixa = " & StatusBar1.Panels(3).Text & ") " & Maquina_Sangria
Set r = dbData.OpenRecordset(SQL)

If r.EOF Then
    Set r = dbData.OpenRecordset(sSQL)
End If


Me.Hide
Set REL_Caixa_Fech_Resumido.ReportMain1.Recordset = r

Dim vStrCaixa As String
Dim vStrCodCaixa As String
vStrCaixa = StatusBar1.Panels(2).Text
vStrCodCaixa = StatusBar1.Panels(3).Text
REL_Caixa_Fech_Resumido.MostrarRetiradas vStrCaixa, vStrCodCaixa



'REL_Caixa_Fech_Resumido.txtDHead.Caption = "FECHAMENTO DE CAIXA - ABERTURA: " & Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")

'REL_Caixa_Fech_Resumido.rfTroco.Caption = Format(txtTotalTroco.Text, "#,##0.00") & " "
'If Not r.EOF Then

'quantidades

'REL_Caixa_Fech_Resumido.rfVendasTotalQuant.Caption = Format(vQuantVendasTotal, 0) & " "


'valores
'Troco
REL_Caixa_Fech_Resumido.rfExtraTroco.Caption = Format(vVlrTroco, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfExtraTrocoQtde.Caption = Format(1, "000") & " "
'vendas
REL_Caixa_Fech_Resumido.rfVendasTotal.Caption = Format(vVlrVendasTotal, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfVendasDinheiro.Caption = Format(vVlrVendasDinheiro, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfVendasPix.Caption = Format(vVlrVendasPix, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfVendasTransferencia.Caption = Format(vVlrVendasTransferencia, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfVendasDeposito.Caption = Format(vVlrVendasDeposito, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfVendasCartao.Caption = Format(vVlrVendasCartao, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfVendasCheque.Caption = Format(vVlrVendasCheque, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfVendasFinanceira.Caption = Format(vVlrVendasFinanceira, ocMONEY) & " "

REL_Caixa_Fech_Resumido.rfVendasQtde.Caption = Format(vQtdeVendasTotal, "000") & " "
REL_Caixa_Fech_Resumido.rfVendasDinheiroQtde.Caption = Format(vQtdeVendasDinheiro, "000") & " "
REL_Caixa_Fech_Resumido.rfVendasPixQtde.Caption = Format(vQtdeVendasPix, "000") & " "
REL_Caixa_Fech_Resumido.rfVendasTransferenciaQtde.Caption = Format(vQtdeVendasTransferencia, "000") & " "
REL_Caixa_Fech_Resumido.rfVendasDepositoQtde.Caption = Format(vQtdeVendasDeposito, "000") & " "
REL_Caixa_Fech_Resumido.rfVendasFinanceiraQtde.Caption = Format(vQtdeVendasFinanceira, "000") & " "
REL_Caixa_Fech_Resumido.rfVendasCartaoQtde.Caption = Format(vQtdeVendasCartao, "000") & " "
REL_Caixa_Fech_Resumido.rfVendasChequeQtde.Caption = Format(vQtdeVendasCheque, "000") & " "

'parcelas
REL_Caixa_Fech_Resumido.rfParcelasTotal.Caption = Format(vVlrParcelasTotal, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfParcelasDinheiro.Caption = Format(vVlrParcelasDinheiro, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfParcelasPix.Caption = Format(vVlrParcelasPix, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfParcelasTransferencia.Caption = Format(vVlrParcelasTransferencia, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfParcelasDeposito.Caption = Format(vVlrParcelasDeposito, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfParcelasCartao.Caption = Format(vVlrParcelasCartao, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfParcelasCheque.Caption = Format(vVlrParcelasCheque, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfParcelasFinanceira.Caption = Format(vVlrParcelasFinanceira, ocMONEY) & " "

REL_Caixa_Fech_Resumido.rfParcelasTotalQtde.Caption = Format(vQtdeParcelasTotal, "000") & " "
REL_Caixa_Fech_Resumido.rfParcelasDinheiroQtde.Caption = Format(vQtdeParcelasDinheiro, "000") & " "
REL_Caixa_Fech_Resumido.rfParcelasPixQtde.Caption = Format(vQtdeParcelasPix, "000") & " "
REL_Caixa_Fech_Resumido.rfParcelasTransferenciaQtde.Caption = Format(vQtdeParcelasTransferencia, "000") & " "
REL_Caixa_Fech_Resumido.rfParcelasDepositoQtde.Caption = Format(vQtdeParcelasDeposito, "000") & " "
REL_Caixa_Fech_Resumido.rfParcelasFinanceiraQtde.Caption = Format(vQtdeParcelasFinanceira, "000") & " "
REL_Caixa_Fech_Resumido.rfParcelasCartaoQtde.Caption = Format(vQtdeParcelasCartao, "000") & " "
REL_Caixa_Fech_Resumido.rfParcelasChequeQtde.Caption = Format(vQtdeParcelasCheque, "000") & " "

'haveres
REL_Caixa_Fech_Resumido.rfHaveresTotal.Caption = Format(vVlrHaveresTotal, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfHaveresDinheiro.Caption = Format(vVlrHaveresDinheiro, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfHaveresPix.Caption = Format(vVlrHaveresPix, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfHaveresTransferencia.Caption = Format(vVlrHaveresTransferencia, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfHaveresDeposito.Caption = Format(vVlrHaveresDeposito, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfHaveresCartao.Caption = Format(vVlrHaveresCartao, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfHaveresCheque.Caption = Format(vVlrHaveresCheque, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfHaveresFinanceira.Caption = Format(vVlrHaveresFinanceira, ocMONEY) & " "

REL_Caixa_Fech_Resumido.rfHaveresTotalQtde.Caption = Format(vQtdeHaveresTotal, "000") & " "
REL_Caixa_Fech_Resumido.rfHaveresDinheiroQtde.Caption = Format(vQtdeHaveresDinheiro, "000") & " "
REL_Caixa_Fech_Resumido.rfHaveresPixQtde.Caption = Format(vQtdeHaveresPix, "000") & " "
REL_Caixa_Fech_Resumido.rfHaveresTransferenciaQtde.Caption = Format(vQtdeHaveresTransferencia, "000") & " "
REL_Caixa_Fech_Resumido.rfHaveresDepositoQtde.Caption = Format(vQtdeHaveresDeposito, "000") & " "
REL_Caixa_Fech_Resumido.rfHaveresFinanceiraQtde.Caption = Format(vQtdeHaveresFinanceira, "000") & " "
REL_Caixa_Fech_Resumido.rfHaveresCartaoQtde.Caption = Format(vQtdeHaveresCartao, "000") & " "
REL_Caixa_Fech_Resumido.rfHaveresChequeQtde.Caption = Format(vQtdeHaveresCheque, "000") & " "

'resumo
Dim vResumoDinheiro As Currency
Dim vResumoPix As Currency
Dim vResumoTransferencia As Currency
Dim vResumoDeposito As Currency
Dim vResumoCartao As Currency
Dim vResumoCheque As Currency
Dim vResumoFinanceira As Currency

vResumoDinheiro = vVlrVendasDinheiro + vVlrParcelasDinheiro + vVlrHaveresDinheiro + vVlrSuprimento
vResumoPix = vVlrVendasPix + vVlrParcelasPix + vVlrHaveresPix
vResumoTransferencia = vVlrVendasTransferencia + vVlrParcelasTransferencia + vVlrHaveresTransferencia
vResumoDeposito = vVlrVendasDeposito + vVlrParcelasDeposito + vVlrHaveresDeposito
vResumoCartao = vVlrVendasCartao + vVlrParcelasCartao + vVlrHaveresCartao
vResumoCheque = vVlrVendasCheque + vVlrParcelasCheque + vVlrHaveresCheque
vResumoFinanceira = vVlrVendasFinanceira + vVlrParcelasFinanceira + vVlrHaveresFinanceira

REL_Caixa_Fech_Resumido.rfResumoDinheiro.Caption = Format(vResumoDinheiro, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfResumoPix.Caption = Format(vResumoPix, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfResumoTransferencia.Caption = Format(vResumoTransferencia, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfResumoDeposito.Caption = Format(vResumoDeposito, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfResumoCartao.Caption = Format(vResumoCartao, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfResumoCheque.Caption = Format(vResumoCheque, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfResumoFinanceira.Caption = Format(vResumoFinanceira, ocMONEY) & " "

Dim vResumoDinheiroQtde As Integer
Dim vResumoPixQtde As Integer
Dim vResumoTransferenciaQtde As Integer
Dim vResumoDepositoQtde As Integer
Dim vResumoCartaoQtde As Integer
Dim vResumoChequeQtde As Integer
Dim vResumoFinanceiraQtde As Integer

vResumoDinheiroQtde = vQtdeVendasDinheiro + vQtdeParcelasDinheiro + vQtdeHaveresDinheiro + vQtdeSuprimento
vResumoPixQtde = vQtdeVendasPix + vQtdeParcelasPix + vQtdeHaveresPix
vResumoTransferenciaQtde = vQtdeVendasTransferencia + vQtdeParcelasTransferencia + vQtdeHaveresTransferencia
vResumoDepositoQtde = vQtdeVendasDeposito + vQtdeParcelasDeposito + vQtdeHaveresDeposito
vResumoFinanceiraQtde = vQtdeVendasFinanceira + vQtdeParcelasFinanceira + vQtdeHaveresFinanceira
vResumoCartaoQtde = vQtdeVendasCartao + vQtdeParcelasCartao + vQtdeHaveresCartao
vResumoChequeQtde = vQtdeVendasCheque + vQtdeParcelasCheque + vQtdeHaveresCheque

REL_Caixa_Fech_Resumido.rfResumoDinheiroQtde.Caption = Format(vResumoDinheiroQtde, "000") & " "
REL_Caixa_Fech_Resumido.rfResumoPixQtde.Caption = Format(vResumoPixQtde, "000") & " "
REL_Caixa_Fech_Resumido.rfResumoTransferenciaQtde.Caption = Format(vResumoTransferenciaQtde, "000") & " "
REL_Caixa_Fech_Resumido.rfResumoDepositoQtde.Caption = Format(vResumoDepositoQtde, "000") & " "
REL_Caixa_Fech_Resumido.rfResumoFinanceiraQtde.Caption = Format(vResumoFinanceiraQtde, "000") & " "
REL_Caixa_Fech_Resumido.rfResumoCartaoQtde.Caption = Format(vResumoCartaoQtde, "000") & " "
REL_Caixa_Fech_Resumido.rfResumoChequeQtde.Caption = Format(vResumoChequeQtde, "000") & " "

'suprimentos
REL_Caixa_Fech_Resumido.rfSuprimentoTotal.Caption = Format(vVlrSuprimento, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfSuprimentoDinheiro.Caption = Format(vVlrSuprimento, ocMONEY) & " "

REL_Caixa_Fech_Resumido.rfSuprimentoTotalQtde.Caption = Format(vQtdeSuprimento, "000") & " "
REL_Caixa_Fech_Resumido.rfSuprimentoDinheiroQtde.Caption = Format(vQtdeSuprimento, "000") & " "


'sangrias
REL_Caixa_Fech_Resumido.rfSangriaTotal.Caption = Format(vVlrSangria, ocMONEY) & " "
REL_Caixa_Fech_Resumido.rfSangriaDinheiro.Caption = Format(vVlrSangria, ocMONEY) & " "

REL_Caixa_Fech_Resumido.rfSangriaTotalQtde.Caption = Format(vQtdeSangria, "000") & " "
REL_Caixa_Fech_Resumido.rfSangriaDinheiroQtde.Caption = Format(vQtdeSangria, "000") & " "

'extra
REL_Caixa_Fech_Resumido.rfeExtraAPrazoQtde.Caption = Format(vQtdeVendasPrazoTotal, "000") & " "
REL_Caixa_Fech_Resumido.rfeExtraAPrazo.Caption = Format(vVlrVendasPrazoTotal, ocMONEY) & " "

REL_Caixa_Fech_Resumido.rfeExtraCanceladasQtde.Caption = Format(vQtdeVendasCanceladoTotal, "000") & " "
REL_Caixa_Fech_Resumido.rfeExtraCanceladas.Caption = Format(vVlrVendasCanceladoTotal, ocMONEY) & " "

REL_Caixa_Fech_Resumido.rfeExtraOrcamentosQtde.Caption = Format(vQtdeVendasOrcamentoTotal, "000") & " "
REL_Caixa_Fech_Resumido.rfeExtraOrcamentos.Caption = Format(vVlrVendasOrcamentoTotal, ocMONEY) & " "

REL_Caixa_Fech_Resumido.rfeExtraConsignadoQtde.Caption = Format(vQtdeVendasConsignadoTotal, "000") & " "
REL_Caixa_Fech_Resumido.rfeExtraConsignado.Caption = Format(vVlrVendasConsignadoTotal, ocMONEY) & " "

REL_Caixa_Fech_Resumido.rfeExtraAluguelQtde.Caption = Format(vQtdeAluguelTotal, "000") & " "
REL_Caixa_Fech_Resumido.rfeExtraAluguel.Caption = Format(vVlrAluguelTotal, ocMONEY) & " "

REL_Caixa_Fech_Resumido.rfeExtraOSQtde.Caption = Format(vQtdeOSTotal, "000") & " "
REL_Caixa_Fech_Resumido.rfeExtraOS.Caption = Format(vVlrOSTotal, ocMONEY) & " "

'REL_Caixa_Fech_Resumido.rfeExtraOSQtde.Caption = Format(vVlrOSTotal, "000") & " "
'REL_Caixa_Fech_Resumido.rfeExtraOS.Caption = Format(vVlrOSTotal, ocMONEY) & " "

'Saldos
Dim vSaldoGeral As Currency
vSaldoGeral = vVlrVendasTotal + vVlrParcelasTotal + vVlrHaveresTotal + vVlrSuprimento
vSaldoGeral = vSaldoGeral - vVlrSangria
REL_Caixa_Fech_Resumido.rfSaldoGeral.Caption = Format(vSaldoGeral, ocMONEY) & " "

Dim vSaldoFisico As Currency
vSaldoFisico = vVlrVendasDinheiro + vVlrParcelasDinheiro + vVlrHaveresDinheiro + vVlrSuprimento
vSaldoFisico = vSaldoFisico + vVlrVendasCheque + vVlrParcelasCheque + vVlrHaveresCheque
vSaldoFisico = vSaldoFisico - vVlrSangria
REL_Caixa_Fech_Resumido.rfSaldoFisico.Caption = Format(vSaldoFisico, ocMONEY) & " "


'===========================RODAPÉ
Dim sSQLusuario As String
Dim r_usuario As ADODB.Recordset

sSQLusuario = "SELECT DATA_ABERTURA, HORA_ABERTURA, COD_FUNC_ABERTURA, DATA_FECHAMENTO, HORA_FECHAMENTO, COD_FUNC_FECHAMENTO, (CASE WHEN status = 1 THEN 'FECHADO' ELSE 'ABERTO' END) AS VarStatus, " & _
        "(SELECT Usuario.Login FROM Usuario INNER JOIN caixa_dia ON Usuario.Codigo = caixa_dia.COD_FUNC_ABERTURA wHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & StatusBar1.Panels(3).Text & ")) AS Nome_Abertura, " & _
        "(SELECT Usuario_2.Login FROM Usuario AS Usuario_2 INNER JOIN caixa_dia AS caixa_dia_2 ON Usuario_2.Codigo = caixa_dia_2.COD_FUNC_FECHAMENTO WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & StatusBar1.Panels(3).Text & ")) AS Nome_Fechamento " & _
       "FROM caixa_dia AS caixa_dia_1 " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and (CODCAIXA = " & StatusBar1.Panels(3).Text & ");"
Set r_usuario = dbData.OpenRecordset(sSQLusuario)


If Not r_usuario.EOF Then
    REL_Caixa_Fech_Resumido.rfCodUsuarioA.Caption = Format(r_usuario("COD_FUNC_ABERTURA"), "00")
    REL_Caixa_Fech_Resumido.rfNomeUsuarioA.Caption = ValidateNull(r_usuario("Nome_Abertura"))
    REL_Caixa_Fech_Resumido.rfDataA.Caption = Format(ValidateNull(r_usuario("DATA_ABERTURA")), "dd/mm/yyyy")
    REL_Caixa_Fech_Resumido.rfHoraA.Caption = Format(ValidateNull(r_usuario("HORA_ABERTURA")), "hh:mm")
    
    REL_Caixa_Fech_Resumido.rfNomeUsuarioF.Caption = ValidateNull(r_usuario("Nome_Fechamento"))
    If IsNull(r_usuario("DATA_FECHAMENTO")) Then
        REL_Caixa_Fech_Resumido.rfDataF.Caption = ""
        REL_Caixa_Fech_Resumido.rfCodUsuarioF.Caption = ""
        REL_Caixa_Fech_Resumido.rfHoraF.Caption = ""
    Else
        REL_Caixa_Fech_Resumido.rfCodUsuarioF.Caption = Format(ValidateNull(r_usuario("COD_FUNC_FECHAMENTO")), "00")
        REL_Caixa_Fech_Resumido.rfDataF.Caption = Format(ValidateNull(r_usuario("DATA_FECHAMENTO")), "dd/mm/yyyy")
        REL_Caixa_Fech_Resumido.rfHoraF.Caption = Format(ValidateNull(r_usuario("HORA_FECHAMENTO")), "hh:mm")
    End If

    REL_Caixa_Fech_Resumido.rfSituacao.Caption = ValidateNull(r_usuario("VARSTATUS"))
End If

REL_Caixa_Fech_Resumido.rfCaixa.Caption = StatusBar1.Panels(2).Text
REL_Caixa_Fech_Resumido.rfCodCaixa.Caption = Format(StatusBar1.Panels(3).Text, "0000")

REL_Caixa_Fech_Resumido.ReportMain1.NomeImpressora = var_ImpNormal
REL_Caixa_Fech_Resumido.ReportMain1.Ativar
'Unload REL_Caixa_Fech_Resumido
Me.Show
End Sub

Private Sub cmdMaqOK_Click()
StatusBar1.Panels(2).Text = cboMaquina.Text
var_Caixa = cboMaquina.Text
If StatusBar1.Panels(2).Text = "TODOS" Then cmdAbrirCaixa.Enabled = False: cmdImprimir.Enabled = True: cmdTroco.Enabled = False Else cmdAbrirCaixa.Enabled = True: cmdImprimir.Enabled = True: cmdTroco.Enabled = True
MostrarCodCaixa
cmdMostrar_Click
frmMaquina.Visible = False
End Sub

Public Sub cmdMostrar_Click()
Dim SETOR_CAIXA As String
'Dim var_Setor As String
Dim varTipoCartao2 As String
Dim sSQL As String
Dim r As ADODB.Recordset

If Not IsDate(mskData) Then Exit Sub

If varCodCaixa = 0 Then
    sSQL = "SELECT " & _
       "parcelas.tipo as varTipoLanc, " & _
       "parcelas.hora AS varHora, " & _
       "parcelas.codigo AS varCodigo, " & _
       "pedidos.cod_pedido AS varCodPedido, " & _
       "cliente.nome AS varCliente, " & _
       "parcelas.forma_pgto AS varFormaPgto, " & _
       "parcelas.valor_final AS varValorLanc, " & _
       "0 AS varValorSaida, " & _
       "(parcelas.valor_final  - 0) AS campo04, " & _
       "(CASE WHEN parcelas.tipo_cartao = 'D' THEN 'DÉBITO' WHEN parcelas.tipo_cartao = 'C'  THEN 'CRÉDITO' Else '' End) AS varTipoCartao, " & _
       "parcelas.pagamento AS data, " & _
       "pedidos.cod_pedido AS pedido, " & _
       "pedidos.cod_cliente AS cliente, " & _
       "pedidos.TIPO_PAGAMENTO AS varTipoPgto, " & _
       "'' AS setor, " & _
       "parcelas.caixa " & _
       "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _
       "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
       "WHERE 1=0"

    Set r = dbData.OpenRecordset(sSQL)
Else
    Dim Maquina_Parcela As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Parcela = "AND (parcelas.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Parcela = "AND (parcelas.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Haver As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Haver = "AND (parcelas_haver.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Haver = "AND (parcelas_haver.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Suprimento As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Suprimento = "AND (caixa_entrada.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Suprimento = "AND (caixa_entrada.caixa <> 'CAIXA') "
    End If
    
    Dim Maquina_Sangria As String
    If StatusBar1.Panels(2).Text <> "TODOS" Then
       Maquina_Sangria = "AND (caixa_saida.caixa = '" & StatusBar1.Panels(2).Text & "') "
    ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
       Maquina_Sangria = "AND (caixa_saida.caixa <> 'CAIXA') "
    End If
    
    'tipo de pedido (balcao, oficina, todos)
    'If StatusBar1.Panels(3).Text <> "TODOS" Then
       SETOR_CAIXA = "AND (pedidos.tipo_pedido = 'VENDA') "
    'ElseIf StatusBar1.Panels(3).Text = "TODOS" Then
    '   SETOR_CAIXA = "AND (pedidos.tipo_pedido <> 'BOSTA') "
    'End If
    
    'setor
    'If StatusBar1.Panels(3).Text <> "TODOS" Then
    '   var_Setor = "AND (setor = '" & StatusBar1.Panels(3).Text & "') "
    'ElseIf StatusBar1.Panels(3).Text = "TODOS" Then
       'var_Setor = "AND (setor <> 'BOSTA') "
    'End If
    
    'Parcelas
    sSQL = "SELECT " & _
       "parcelas.tipo as varTipoLanc, " & _
       "parcelas.hora AS varHora, " & _
       "parcelas.codigo AS varCodigo, " & _
       "pedidos.cod_pedido AS varCodPedido, " & _
       "cliente.nome AS varCliente, " & _
       "parcelas.forma_pgto AS varFormaPgto, " & _
       "parcelas.valor_final AS varValorLanc, " & _
       "0 AS varValorSaida, " & _
       "(parcelas.valor_final  - 0) AS campo04, " & _
       "(CASE WHEN parcelas.tipo_cartao = 'D' THEN 'DÉBITO' WHEN parcelas.tipo_cartao = 'C'  THEN 'CRÉDITO' Else '' End) AS varTipoCartao, " & _
       "parcelas.pagamento AS data, " & _
       "pedidos.cod_pedido AS pedido, " & _
       "pedidos.cod_cliente AS cliente, " & _
       "pedidos.TIPO_PAGAMENTO AS varTipoPgto, " & _
       "'' AS setor, " & _
       "parcelas.caixa " & _
       "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _
       "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
       "WHERE (parcelas.status = 1) AND (parcelas.codcaixa = " & StatusBar1.Panels(3).Text & ") " & Maquina_Parcela & _
       "UNION ALL "
    
    'CASE WHEN pedidos.tipo_cartao = 'D' THEN 'DÉBITO' WHEN pedidos.tipo_cartao = 'C'  THEN 'CRÉDITO' Else '' End
    'parcelas_haveres
    sSQL = sSQL & "SELECT " & _
       "'HAVER' as varTipoLanc, " & _
       "parcelas_haver.hora AS varHora, " & _
       "0 AS varCodigo, " & _
       "parcelas.cod_pedido AS varCodPedido, " & _
       "(SELECT nome FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente WHERE (pedidos.cod_pedido = parcelas.cod_pedido)) AS varCliente, " & _
       "parcelas_haver.forma_pgto AS varFormaPgto, " & _
       "parcelas_haver.valor_haver AS varValorLanc, " & _
       "0 AS varValorSaida, " & _
       "parcelas_haver.valor_haver AS campo04, " & _
       "(CASE WHEN parcelas_haver.tipo_cartao = 'D' THEN 'DÉBITO' WHEN parcelas_haver.tipo_cartao = 'C'  THEN 'CRÉDITO' Else '' End) AS varTipoCartao, " & _
       "parcelas_haver.haver AS data, " & _
       "parcelas_haver.codigo AS campotc, " & _
       "0 AS cliente, " & _
       "'' AS varTipoPgto, " & _
       "'' AS  setor, " & _
       "'' AS maquina " & _
       "FROM parcelas_haver INNER JOIN parcelas ON parcelas_haver.cod_parcela = parcelas.codigo " & _
       "WHERE (parcelas_haver.codcaixa = " & StatusBar1.Panels(3).Text & ") " & Maquina_Haver & _
       "UNION ALL "

    'suprimentos
    sSQL = sSQL & "SELECT " & _
       "'SUPRIMENTO' as varTipoLanc, " & _
       "hora AS varHora, " & _
       "0 AS varCodigo, " & _
       "0 AS varCodPedido, " & _
       "descricao AS varCliente, " & _
       "caixa_entrada.forma_pgto AS varFormaPgto, " & _
       "valor AS varValorLanc, " & _
       "0 AS varValorSaida, " & _
       "valor AS campo04, " & _
       "'' AS varTipoCartao, " & _
       "data, " & _
       "0 AS pedido, " & _
       "0 AS cliente, " & _
       "'' AS varTipoPgto, " & _
       "setor, " & _
       " '' AS maquina " & _
       "FROM caixa_entrada WHERE (caixa_entrada.codcaixa = " & StatusBar1.Panels(3).Text & ")  " & Maquina_Suprimento & _
       "UNION ALL "

    'sangria
    sSQL = sSQL & "SELECT " & _
       "'SANGRIA' as varTipoLanc, " & _
       "caixa_saida.hora AS varHora, " & _
       "0 AS varCodigo, " & _
       "0 AS varCodPedido, " & _
       "caixa_saida.descricao AS varCliente, " & _
       "'DINHEIRO' AS varFormaPgto, " & _
       "0 AS varValorLanc, " & _
       "caixa_saida.valor AS varValorSaida, " & _
       "(0 - caixa_saida.valor) AS campo04, " & _
       "'' AS varTipoCartao, " & _
       "data, " & _
       "0, " & _
       "0, " & _
       "'' AS varTipoPgto, " & _
       "setor, " & _
       "'' AS maquina " & _
       "FROM caixa_saida WHERE (FONTE = 'CAIXA ATUAL') AND (caixa_saida.codcaixa = " & StatusBar1.Panels(3).Text & ")  " & Maquina_Sangria & _
       " ORDER BY 2"
'Debug.Print sSQL
    Set r = dbData.OpenRecordset(sSQL)
End If

If r.RecordCount = 0 Then cmdDetalhar.Enabled = False Else cmdDetalhar.Enabled = True

'If r("vartipocartao") = "D" Then
'     varTipoCartao2 = r("varFormaPgto") & " DÉBITO"
' Else
'     varTipoCartao2 = r("varFormaPgto") & " CRÉDITO"
' End If
Mostrar_APrazo
Mostrar_Aluguel
Mostrar_Servico
Mostrar_Retiradas

FormatarGridEntrada r
CompararCaixa
  
If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL

'MOSTRAR As ENTRADAS

End Sub
Private Sub cmdOKData_Click()
StatusBar1.Panels(5).Text = Format(mskData, "dd/mm/yyyy")
frmData.Visible = False
cmdMostrar_Click
Form_Activate
End Sub

Private Sub cmdOKSetor_Click()
'StatusBar1.Panels(3).Text = cboSetor.Text
'cmdMostrar_Click
'frmSetor.Visible = False
End Sub

Private Sub cmdSalvarTroco_Click()
Dim x_Troco As Long

If txtTroco.Text = "" Then frmTroco.Visible = False: Exit Sub

'CHECAR SE O JÁ TEM TROCO ADICIONADO PARA A DATA
sSQL = "SELECT * FROM caixa_troco WHERE (caixa_troco.codcaixa = " & varCodCaixa & ") AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   dbData.Execute "UPDATE caixa_troco SET valor = " & Replace(CCur(txtTroco.Text), ",", ".") & ", caixa = '" & StatusBar1.Panels(2).Text & "', codcaixa = " & varCodCaixa & " WHERE codcaixa = " & varCodCaixa & " AND (caixa = '" & StatusBar1.Panels(2).Text & "') ;"
Else
   x_Troco = 1
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_troco FROM caixa_troco where (caixa = '" & StatusBar1.Panels(2).Text & "');"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then x_Troco = r("ultimo_troco") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   dbData.Execute "INSERT INTO caixa_troco (codigo, data, valor, caixa, codcaixa) VALUES (" & x_Troco & ", CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103), " & Replace(CCur(txtTroco.Text), ",", ".") & ", '" & StatusBar1.Panels(2).Text & "', " & varCodCaixa & ");"
End If

txtTroco.Text = ""
frmTroco.Visible = False
'lblAviso1.Visible = True
cmdMostrar_Click
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

Private Sub cmdSenha_Click()
If var_Caixa = "TODOS" Then
    MsgBox "Escolha qual o nome do caixa correto!", vbInformation, "aviso do sistema"
    Exit Sub
End If

sSQL = "SELECT * FROM usuario WHERE (password = '" & txtSenha.Text & "');"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   txtSenha.Text = ""
   frmSenha.Visible = False
   varBotaoCaixa = True
   Caixa_Fechamento.Show 1
Else
   ShowMsg "ACESSO NEGADO!" & vbCrLf & "Você não tem nivel de acesso a esse recurso", vbInformation
   txtSenha.Text = ""
   frmSenha.Visible = False
   Exit Sub
End If
End Sub

Private Sub cmdTrocarCaixa_Click()
If lblCodCaixaAtual.Caption = "0" Or lblCodCaixaStatus.Caption = "FECHADO" Then MsgBox "O " & lblCaixaAtual.Caption & " ainda encontra fechado!", vbInformation, "Aviso do Sistema": Exit Sub

i = Grid.Row

If Grid.TextMatrix(i, 3) <> "VENDA" Then MsgBox "Somente é possível a troca de caixa para VENDAS!", vbInformation, "Aviso do Sistema": Exit Sub

If ShowMsg("Tem certeza que deseja mudar de caixa a venda de " & Grid.TextMatrix(i, 4) & " no valor de " & Format(Grid.TextMatrix(i, 6), ocMONEY) & " ?", vbInformation + vbYesNo) = vbYes Then
    sSQL = "SELECT caixa, codcaixa " & _
           "FROM caixa_dia " & _
           "WHERE (caixa = '" & lblCaixaAtual & "') and (codcaixa = " & lblCodCaixaAtual & ");"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.EOF Then
        dbData.Execute "UPDATE pedidos SET caixa = '" & lblCaixaAtual & "', codcaixa = " & lblCodCaixaAtual & " WHERE COD_PEDIDO = " & Val(Grid.TextMatrix(i, 2)) & "  ;"
        dbData.Execute "UPDATE parcelas SET caixa = '" & lblCaixaAtual & "', codcaixa = " & lblCodCaixaAtual & " WHERE COD_PEDIDO = " & Val(Grid.TextMatrix(i, 2)) & "  ;"
        Call cmdMostrar_Click
    Else
        MsgBox "Transferência de caixa incorreta!", vbInformation, "Aviso do Sistema"
    End If
End If
End Sub

Private Sub cmdTroco_Click()
frmTroco.Visible = True
'lblAviso1.Visible = False

sSQL = "SELECT * FROM caixa_troco WHERE (caixa_troco.codcaixa = " & varCodCaixa & ") AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then txtTroco.Text = Format(r("valor"), ocMONEY)
If r.State <> 0 Then r.Close
Set r = Nothing

txtTroco.SetFocus
End Sub

Private Sub Form_Activate()
VerificarCaixa
End Sub

Private Sub Form_Load()
Dim oIni As Ini
varTipoConsulta = ""

'mostrar os objetos de OS e/ou Aluguel
Dim oCfg As ConfigItem
Dim bStatus As Boolean

'os
Set oCfg = sysConfig("os")    'Recupera a config deseja
bStatus = CBool(oCfg.Value)   'Converte o valor para booleano
vOSAtiva = CBool(oCfg.Value)
Set oCfg = Nothing            'Destroi o objeto

txtQuantDinheiroOS.Visible = bStatus 'Habilita/desabilida conforme valor
txtTotalDinheiroOS.Visible = bStatus
lblOS.Visible = bStatus

'aluguel
Set oCfg = sysConfig("aluguel")    'Recupera a config deseja
bStatus = CBool(oCfg.Value)   'Converte o valor para booleano
vAluguelAtiva = CBool(oCfg.Value)
'Set oCfg = Nothing            'Destroi o objeto

txtQuantDinheiroAluguel.Visible = bStatus 'Habilita/desabilida conforme valor
txtTotalDinheiroAluguel.Visible = bStatus
lblAluguel.Visible = bStatus
'Set oCfg = Nothing

'tipo de caixa
Set oCfg = sysConfig("TIPOCAIXA")
vTipoCaixa = CInt(oCfg.Value)
Set oCfg = Nothing

If vTipoCaixa = 2 Then
    cmdTrocarCaixa.Visible = True
    cmdCaixaPrincipal.Visible = True
    cmdDetalhar.Visible = True
Else
    cmdTrocarCaixa.Visible = False
    cmdCaixaPrincipal.Visible = False
    cmdDetalhar.Visible = True
End If

EsconderTotais

If varFluxoCaixa = False Then
    Set oIni = New Ini
    oIni.Arquivo = appPathApp & "config.ini"
    'var_Setor = oIni.LerTexto("DADOS_SETOR", "setor")
    'var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
    var_Caixa = oIni.LerTexto("DADOS_CAIXA", "caixa")
    Set oIni = Nothing
    
    StatusBar1.Panels(2).Text = var_Caixa
    lblCaixaAtual.Caption = var_Caixa
    CompararCaixa
    cmdCaixaPrincipal.Caption = "Caixa Principal"
    lblCaixaRotulo.Caption = var_Caixa
    'StatusBar1.Panels(5).Text = Format(Date - 1, "dd/mm/yy")
    'StatusBar1.Panels(2).Text = var_Maquina
    StatusBar1.Panels(5).Text = Format(Date, "dd/mm/yyyy")
    mskData.Text = Format(Date, "dd/mm/yyyy")
    frmTroco.Visible = False
    frmMaquina.Visible = False
    cmdAbrirCaixa.Visible = False
    cmdFecharCaixa.Visible = False
    
    MostrarCodCaixa
    cmdMostrar_Click
    varBotaoCaixa = False
    lblCodCaixaAtual.Caption = varCodCaixa
    lblCodCaixaStatus.Caption = vStatusCaixaAtual
Else
    StatusBar1.Panels(2).Text = varFluxoNomeCaixa
    StatusBar1.Panels(3).Text = varFluxoCodCaixa
    StatusBar1.Panels(4).Text = varFluxoCaixaSituacao
    StatusBar1.Panels(5).Text = Format(varFluxoCaixaData, "dd/mm/yyyy")
    mskData.Text = Format(varFluxoCaixaData, "dd/mm/yyyy")
    varCodCaixa = varFluxoCodCaixa
    cmdAbrirCaixa.Enabled = False
    cmdFecharCaixa.Enabled = False
    cmdTroco.Enabled = False
    cmdImprimir.Enabled = True
    cmdImprimirResumido.Enabled = True
    cmdMostrar_Click
    lblCodCaixaAtual.Caption = varCodCaixa
    lblCodCaixaStatus.Caption = vStatusCaixaAtual
End If

Set moCombo = New cComboHelper
End Sub
Private Sub MostrarCodCaixa()
sSQL = "SELECT *, CASE status WHEN 0 THEN 'ABERTO' ELSE 'FECHADO' END AS varStatus " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    varCodCaixa = ValidateNull(r("codcaixa"))
    cmdTroco.Enabled = True
    cmdImprimir.Enabled = True
    cmdAbrirCaixa.Visible = False
    cmdFecharCaixa.Visible = True
    StatusBar1.Panels(3).Text = Format(ValidateNull(r("codcaixa")), "00000")
    StatusBar1.Panels(4).Text = r("VARSTATUS")
Else
    If varFluxoCaixa = False Then
        varCodCaixa = 0
        cmdTroco.Enabled = False
        cmdImprimir.Enabled = False
        cmdAbrirCaixa.Visible = True
        cmdFecharCaixa.Visible = False
        StatusBar1.Panels(3).Text = Format(0, "00000")
        StatusBar1.Panels(4).Text = ""
    Else
        varCodCaixa = StatusBar1.Panels(3).Text
    End If
End If
End Sub

Private Sub FormatarGridEntradaDetalhado(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   Dim m_Saldo As Currency
   
   With Grid
      .Clear
      .Cols = 7
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 650
      .ColWidth(2) = 4800
      .ColWidth(3) = 1150
      .ColWidth(4) = 1150
      .ColWidth(5) = 1150
      .ColWidth(6) = 1150
      
      .TextMatrix(0, 1) = "HORA"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "TIPO"
      .TextMatrix(0, 4) = "ENTRADA"
      .TextMatrix(0, 5) = "SAÍDA"
      .TextMatrix(0, 6) = "SALDO"
      
      .Row = 0
      .Redraw = False
      
      'colocar os cabeçalho em negrito / Centralizado
      For i = 0 To .Cols - 1
         .Col = i
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      i = 1
      m_Saldo = 0
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("varHora"), ocHRMN)
            .TextMatrix(.rows - 1, 2) = rTabela("varCliente")
            .TextMatrix(.rows - 1, 3) = rTabela("var_tipo")
            .TextMatrix(.rows - 1, 4) = Format(rTabela("varValorLanc"), ocMONEY)
            .TextMatrix(.rows - 1, 5) = Format(rTabela("varValorSaida"), ocMONEY)
            
            m_Saldo = m_Saldo + CCur(rTabela("varValorLanc")) - CCur(rTabela("varValorSaida"))
            .TextMatrix(.rows - 1, 5) = Format(m_Saldo, ocMONEY)
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .rows = .rows - 1
      
      'mudar a cor da coluna
      For i = 1 To .rows - 1
         .Row = i
         .Col = 4:   .CellBackColor = &HC0FFFF
         .Col = 5:   .CellBackColor = &HC0C0FF
      Next

      'Deixar negrito quando vencido
      For i = 1 To .rows - 1
         For j = 0 To .Cols - 1
            .Col = j
            .Row = i
            If CCur(.TextMatrix(i, 4)) > 0 Then .CellFontBold = True
         Next
      Next
      
      .Redraw = True
   End With
End Sub

Private Sub FormatarGridResumido(rTabela As ADODB.Recordset)
   Dim i As Integer
'   Dim m_Saldo As Currency
   
   With Grid
      .Clear
      .Cols = 3
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 3000
      .ColWidth(2) = 1000

      .TextMatrix(0, 1) = "DESCRIÇÃO"
      .TextMatrix(0, 2) = "VALOR"
      
      .Row = 0
      
      'colocar os cabeçalho em negrito / Centralizado
      For i = 0 To .Cols - 1
         .Col = i
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      'm_Saldo = 0
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = ValidateNull(rTabela("vTipoResultado"))
            .TextMatrix(.rows - 1, 2) = Format(rTabela("vtotal"), ocMONEY)
           
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .rows = .rows - 1
      
      'Deixar negrito quando vencido
'      For i = 1 To .Rows - 1
'         For j = 0 To .Cols - 1
'            .Col = j
'            .Row = i
            
'            If Left(.TextMatrix(i, 5), 6) = "CARTAO" Or .TextMatrix(i, 5) = "PIX" Then
'               .CellForeColor = &H8000&
'               .CellFontBold = True
'            ElseIf .TextMatrix(i, 5) = "DINHEIRO" And .TextMatrix(i, 3) <> "SANGRIA" Then
'               txtTotalDinheiro.Text = Format(SomaGrid(Grid, 6), ocMONEY)
'               .CellForeColor = vbBlack
'            ElseIf .TextMatrix(i, 3) = "SANGRIA" And .TextMatrix(i, 5) = "DINHEIRO" Then
'               .CellForeColor = vbRed
'               .CellFontBold = True
'            End If
'         Next
'      Next
      
      .Redraw = True
   End With
   
'SomaFlexDinheiro
'SomaFlexCartao
'SomaFlexCheque
'SomaFlexOutros
'SomaFlexSaida
'SomaFaturamento
'Mostrar_Troco
'Mostrar_Saldo
End Sub

Private Sub FormatarGridEntrada(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   Dim m_Saldo As Currency
   
   With Grid
      .Clear
      .Cols = 11
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 650
      .ColWidth(2) = 850
      .ColWidth(3) = 1300
      .ColWidth(4) = 4700
      .ColWidth(5) = 2000
      .ColWidth(6) = 1050
      .ColWidth(7) = 1050
      .ColWidth(8) = 1050
      .ColWidth(9) = 0
      .ColWidth(10) = 0

      
      .TextMatrix(0, 1) = "HORA"
      .TextMatrix(0, 2) = "PEDIDO"
      .TextMatrix(0, 3) = "TIPO"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "FORMA"
      .TextMatrix(0, 6) = "ENTRADA"
      .TextMatrix(0, 7) = "SAÍDA"
      .TextMatrix(0, 8) = "SALDO"
      .TextMatrix(0, 9) = "COD_PARC"
      .TextMatrix(0, 10) = "TIPO_PGTO"
      
      .Row = 0
      
      'colocar os cabeçalho em negrito / Centralizado
      For i = 0 To .Cols - 1
         .Col = i
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      m_Saldo = 0
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("varHora"), ocHRMN)
            If rTabela("varCodPedido") = 0 Then
                .TextMatrix(.rows - 1, 2) = ""
            Else
                .TextMatrix(.rows - 1, 2) = Format(rTabela("varCodPedido"), "000000")
            End If
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("varTipoLanc"))
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("varCliente"))
            
            If rTabela("varFormaPgto") <> "CARTAO" Then
               .TextMatrix(.rows - 1, 5) = rTabela("varFormaPgto")
            Else
                If rTabela("vartipocartao") = "DÉBITO" Then
                    .TextMatrix(.rows - 1, 5) = rTabela("varFormaPgto") & " DÉBITO"
                Else
                    .TextMatrix(.rows - 1, 5) = rTabela("varFormaPgto") & " CRÉDITO"
                End If

            '.TextMatrix(.Rows - 1, 5) = rTabela("varFormaPgto") & " (" & rTabela("vartipocartao") & ")"
               
               
            End If
            
            .TextMatrix(.rows - 1, 6) = Format(rTabela("varValorLanc"), ocMONEY)
            .TextMatrix(.rows - 1, 7) = Format(rTabela("varValorSaida"), ocMONEY)
            
            m_Saldo = m_Saldo + CCur(ValidateNull(rTabela("varValorLanc"))) - CCur(rTabela("varValorSaida"))
            .TextMatrix(.rows - 1, 8) = Format(m_Saldo, "##,##0.00")
            .TextMatrix(.rows - 1, 9) = Format(rTabela("varCodigo"), "000000")
            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("varTipoPgto"))
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .rows = .rows - 1
      
      'mudar a cor da coluna
      For i = 1 To .rows - 1
         .Row = i
         .Col = 6:   .CellBackColor = &HC0FFFF
         .Col = 7:   .CellBackColor = &HC0C0FF
      Next
      
      'Deixar negrito quando vencido
      For i = 1 To .rows - 1
         For j = 0 To .Cols - 1
            .Col = j
            .Row = i
            
            If Left(.TextMatrix(i, 5), 6) = "CARTAO" Or .TextMatrix(i, 5) = "PIX" Then
               .CellForeColor = &H8000&
               .CellFontBold = True
            ElseIf .TextMatrix(i, 5) = "DINHEIRO" And .TextMatrix(i, 3) <> "SANGRIA" Then
               txtTotalDinheiro.Text = Format(SomaGrid(Grid, 6), ocMONEY)
               .CellForeColor = vbBlack
            ElseIf .TextMatrix(i, 3) = "SANGRIA" And .TextMatrix(i, 5) = "DINHEIRO" Then
               .CellForeColor = vbRed
               .CellFontBold = True
            End If
         Next
      Next
      
      .Redraw = True
   End With
   
SomaFlexDinheiro
SomaFlexCartao
SomaFlexCheque
SomaFlexOutros
SomaFlexSaida
SomaFaturamento
Mostrar_Troco
Mostrar_Saldo
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'SABER O NOME DO PROJETO ABERTO
'If varBotaoCaixa = True Then
    'Unload Principal_Caixa
    ''Unload PDV
    'Unload Me
'Else
    ''HabilitaObjetosVenda False
'End If
If vChamouCaixa = "PDV" Then
    Caixa_Controle_semOS.Hide
    'PDV.Show  'desativei somente para geerar o online comerce
Else
    Caixa_Controle_semOS.Hide
    'PDV.Show 1
End If

varFluxoCaixa = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
If cmdTrocarCaixa.Enabled = True Then
    Call cmdTrocarCaixa_Click
End If
End Sub

Private Sub img1_Click()
'Me.Hide
varFluxoNomeCaixa = StatusBar1.Panels(2).Text
varFluxoCodCaixa = StatusBar1.Panels(3).Text
varTipoConsulta = "VENDAS"
Caixa_Controle_Resumo.StatusBar1.Panels(2).Text = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
Caixa_Controle_Resumo.StatusBar1.Panels(3).Text = Caixa_Controle_semOS.StatusBar1.Panels(3).Text
Caixa_Controle_Resumo.StatusBar1.Panels(4).Text = Caixa_Controle_semOS.StatusBar1.Panels(4).Text
Caixa_Controle_Resumo.StatusBar1.Panels(5).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
Caixa_Controle_Resumo.Show
End Sub

Private Sub img10_Click()
varFluxoNomeCaixa = StatusBar1.Panels(2).Text
varFluxoCodCaixa = StatusBar1.Panels(3).Text
varTipoConsulta = "SERVICOS"
Caixa_Controle_Resumo.StatusBar1.Panels(2).Text = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
Caixa_Controle_Resumo.StatusBar1.Panels(3).Text = Caixa_Controle_semOS.StatusBar1.Panels(3).Text
Caixa_Controle_Resumo.StatusBar1.Panels(4).Text = Caixa_Controle_semOS.StatusBar1.Panels(4).Text
Caixa_Controle_Resumo.StatusBar1.Panels(5).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
Caixa_Controle_Resumo.Show
End Sub

Private Sub img11_Click()
'Me.Hide
varFluxoNomeCaixa = StatusBar1.Panels(2).Text
varFluxoCodCaixa = StatusBar1.Panels(3).Text
varTipoConsulta = "ALUGUEL"
Caixa_Controle_Resumo.StatusBar1.Panels(2).Text = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
Caixa_Controle_Resumo.StatusBar1.Panels(3).Text = Caixa_Controle_semOS.StatusBar1.Panels(3).Text
Caixa_Controle_Resumo.StatusBar1.Panels(4).Text = Caixa_Controle_semOS.StatusBar1.Panels(4).Text
Caixa_Controle_Resumo.StatusBar1.Panels(5).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
Caixa_Controle_Resumo.Show
End Sub


Private Sub img2_Click()
'Me.Hide
varFluxoNomeCaixa = StatusBar1.Panels(2).Text
varFluxoCodCaixa = StatusBar1.Panels(3).Text
varTipoConsulta = "PARCELAS"
Caixa_Controle_Resumo.StatusBar1.Panels(2).Text = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
Caixa_Controle_Resumo.StatusBar1.Panels(3).Text = Caixa_Controle_semOS.StatusBar1.Panels(3).Text
Caixa_Controle_Resumo.StatusBar1.Panels(4).Text = Caixa_Controle_semOS.StatusBar1.Panels(4).Text
Caixa_Controle_Resumo.StatusBar1.Panels(5).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
Caixa_Controle_Resumo.Show
End Sub


Private Sub img3_Click()
'Me.Hide
varFluxoNomeCaixa = StatusBar1.Panels(2).Text
varFluxoCodCaixa = StatusBar1.Panels(3).Text
varTipoConsulta = "HAVERES"
Caixa_Controle_Resumo.StatusBar1.Panels(2).Text = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
Caixa_Controle_Resumo.StatusBar1.Panels(3).Text = Caixa_Controle_semOS.StatusBar1.Panels(3).Text
Caixa_Controle_Resumo.StatusBar1.Panels(4).Text = Caixa_Controle_semOS.StatusBar1.Panels(4).Text
Caixa_Controle_Resumo.StatusBar1.Panels(5).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
Caixa_Controle_Resumo.Show
End Sub


Private Sub img4_Click()
'Me.Hide
varFluxoNomeCaixa = StatusBar1.Panels(2).Text
varFluxoCodCaixa = StatusBar1.Panels(3).Text
varTipoConsulta = "SUPRIMENTOS"
Caixa_Controle_Resumo.StatusBar1.Panels(2).Text = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
Caixa_Controle_Resumo.StatusBar1.Panels(3).Text = Caixa_Controle_semOS.StatusBar1.Panels(3).Text
Caixa_Controle_Resumo.StatusBar1.Panels(4).Text = Caixa_Controle_semOS.StatusBar1.Panels(4).Text
Caixa_Controle_Resumo.StatusBar1.Panels(5).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
Caixa_Controle_Resumo.Show
End Sub


Private Sub img5_Click()
'Me.Hide
varFluxoNomeCaixa = StatusBar1.Panels(2).Text
varFluxoCodCaixa = StatusBar1.Panels(3).Text
varTipoConsulta = "PRAZO"
Caixa_Controle_Resumo.StatusBar1.Panels(2).Text = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
Caixa_Controle_Resumo.StatusBar1.Panels(3).Text = Caixa_Controle_semOS.StatusBar1.Panels(3).Text
Caixa_Controle_Resumo.StatusBar1.Panels(4).Text = Caixa_Controle_semOS.StatusBar1.Panels(4).Text
Caixa_Controle_Resumo.StatusBar1.Panels(5).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
Caixa_Controle_Resumo.Show
End Sub

Private Sub img6_Click()
'Me.Hide
varFluxoNomeCaixa = StatusBar1.Panels(2).Text
varFluxoCodCaixa = StatusBar1.Panels(3).Text
varTipoConsulta = "SANGRIAS"
Caixa_Controle_Resumo.StatusBar1.Panels(2).Text = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
Caixa_Controle_Resumo.StatusBar1.Panels(3).Text = Caixa_Controle_semOS.StatusBar1.Panels(3).Text
Caixa_Controle_Resumo.StatusBar1.Panels(4).Text = Caixa_Controle_semOS.StatusBar1.Panels(4).Text
Caixa_Controle_Resumo.StatusBar1.Panels(5).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
Caixa_Controle_Resumo.Show
End Sub

Private Sub img7_Click()
'Me.Hide
varFluxoNomeCaixa = StatusBar1.Panels(2).Text
varFluxoCodCaixa = StatusBar1.Panels(3).Text
varTipoConsulta = "CARTAO"
Caixa_Controle_Resumo.StatusBar1.Panels(2).Text = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
Caixa_Controle_Resumo.StatusBar1.Panels(3).Text = Caixa_Controle_semOS.StatusBar1.Panels(3).Text
Caixa_Controle_Resumo.StatusBar1.Panels(4).Text = Caixa_Controle_semOS.StatusBar1.Panels(4).Text
Caixa_Controle_Resumo.StatusBar1.Panels(5).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
Caixa_Controle_Resumo.Show
End Sub


Private Sub img8_Click()
'Me.Hide
varFluxoNomeCaixa = StatusBar1.Panels(2).Text
varFluxoCodCaixa = StatusBar1.Panels(3).Text
varTipoConsulta = "OUTROS"
Caixa_Controle_Resumo.StatusBar1.Panels(2).Text = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
Caixa_Controle_Resumo.StatusBar1.Panels(3).Text = Caixa_Controle_semOS.StatusBar1.Panels(3).Text
Caixa_Controle_Resumo.StatusBar1.Panels(4).Text = Caixa_Controle_semOS.StatusBar1.Panels(4).Text
Caixa_Controle_Resumo.StatusBar1.Panels(5).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
Caixa_Controle_Resumo.Show
End Sub


Private Sub img9_Click()
'Me.Hide
varFluxoNomeCaixa = StatusBar1.Panels(2).Text
varFluxoCodCaixa = StatusBar1.Panels(3).Text
varTipoConsulta = "RETIRADAS"
Caixa_Controle_Resumo.StatusBar1.Panels(2).Text = Caixa_Controle_semOS.StatusBar1.Panels(2).Text
Caixa_Controle_Resumo.StatusBar1.Panels(3).Text = Caixa_Controle_semOS.StatusBar1.Panels(3).Text
Caixa_Controle_Resumo.StatusBar1.Panels(4).Text = Caixa_Controle_semOS.StatusBar1.Panels(4).Text
Caixa_Controle_Resumo.StatusBar1.Panels(5).Text = Format(Caixa_Controle_semOS.StatusBar1.Panels(5).Text, "dd/mm/yyyy")
Caixa_Controle_Resumo.Show
End Sub

Private Sub mskData_GotFocus()
   SelectControl mskData
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
   mskData.Mask = "##/##/##"
End Sub

Private Sub mskData_LostFocus()
   If Not IsDate(mskData.Text) Then
      mskData.Mask = ""
      mskData.Text = ""
   End If
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
   Select Case Panel.index
      Case 1
         Exit Sub
      Case 2
         frmMaquina.Visible = True
         cboMaquina.SetFocus
      Case 3
         'frmSetor.Visible = True
         'cboSetor.SetFocus
         Exit Sub
      Case 4
         Exit Sub
      Case 5
         frmData.Visible = True
         mskData.SetFocus
   End Select
End Sub

Private Sub txtFATTotalAluguel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
img1.Visible = False
img2.Visible = False
img3.Visible = False
img4.Visible = False
img5.Visible = False
img6.Visible = False
img7.Visible = False
img8.Visible = False
img9.Visible = False
img10.Visible = False
img11.Visible = True
End Sub


Private Sub txtFATTotalHaveres_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
img1.Visible = False
img2.Visible = False
img3.Visible = True
img4.Visible = False
img5.Visible = False
img6.Visible = False
img7.Visible = False
img8.Visible = False
img9.Visible = False
img10.Visible = False
img11.Visible = False
End Sub


Private Sub txtFATTotalParcelas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
img1.Visible = False
img2.Visible = True
img3.Visible = False
img4.Visible = False
img5.Visible = False
img6.Visible = False
img7.Visible = False
img8.Visible = False
img9.Visible = False
img10.Visible = False
img11.Visible = False
End Sub


Private Sub txtFATTotalPrazo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
img1.Visible = False
img2.Visible = False
img3.Visible = False
img4.Visible = False
img5.Visible = True
img6.Visible = False
img7.Visible = False
img8.Visible = False
img9.Visible = False
img10.Visible = False
img11.Visible = False
End Sub


Private Sub txtFATTotalSaidas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
img1.Visible = False
img2.Visible = False
img3.Visible = False
img4.Visible = False
img5.Visible = False
img6.Visible = True
img7.Visible = False
img8.Visible = False
img9.Visible = False
img10.Visible = False
img11.Visible = False
End Sub


Private Sub txtFATTotalServicos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
img1.Visible = False
img2.Visible = False
img3.Visible = False
img4.Visible = False
img5.Visible = False
img6.Visible = False
img7.Visible = False
img8.Visible = False
img9.Visible = False
img10.Visible = True
img11.Visible = False
End Sub


Private Sub txtFATTotalSuprimentos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
img1.Visible = False
img2.Visible = False
img3.Visible = False
img4.Visible = True
img5.Visible = False
img6.Visible = False
img7.Visible = False
img8.Visible = False
img9.Visible = False
img10.Visible = False
img11.Visible = False
End Sub


Private Sub txtFATTotalVendas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
img1.Visible = True
img2.Visible = False
img3.Visible = False
img4.Visible = False
img5.Visible = False
img6.Visible = False
img7.Visible = False
img8.Visible = False
img9.Visible = False
img10.Visible = False
img11.Visible = False
End Sub


Private Sub txtSenha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSenha_Click
End Sub

Private Sub txtTotalAvulso_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
img1.Visible = False
img2.Visible = False
img3.Visible = False
img4.Visible = False
img5.Visible = False
img6.Visible = False
img7.Visible = False
img8.Visible = True
img8.Top = txtTotalAvulso.Top
img9.Visible = False
End Sub


Private Sub txtTotalCartao_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
img1.Visible = False
img2.Visible = False
img3.Visible = False
img4.Visible = False
img5.Visible = False
img6.Visible = False
img7.Visible = True
img7.Top = txtTotalCartao.Top
img8.Visible = False
img9.Visible = False
End Sub


Private Sub txtTotalRetiradas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
img1.Visible = False
img2.Visible = False
img3.Visible = False
img4.Visible = False
img5.Visible = False
img6.Visible = False
img7.Visible = False
img8.Visible = False
img9.Visible = True
img9.Top = txtTotalRetiradas.Top
End Sub


Private Sub txtTroco_GotFocus()
SelectControl txtTroco
End Sub
