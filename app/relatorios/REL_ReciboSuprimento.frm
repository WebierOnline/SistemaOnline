VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_ReciboSuprimento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão - Recibo"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12075
   Icon            =   "REL_ReciboSuprimento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   110.861
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   212.99
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   60
      TabIndex        =   0
      Top             =   5820
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   10
      Regua           =   -1  'True
      Titulo          =   ""
      Registrado      =   0   'False
      Visualizar      =   0   'False
      Resolucao       =   -1
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   5850
      Left            =   0
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   10319
      Ordem           =   1
      Begin ReportX.ReportField txtFormaPgto 
         Height          =   210
         Left            =   1620
         TabIndex        =   18
         Top             =   5460
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   370
         Caption         =   "ASSINATURA"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUsuario 
         Height          =   210
         Left            =   1320
         TabIndex        =   17
         Top             =   5220
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   370
         Caption         =   "ASSINATURA"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField2 
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   5460
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   370
         Caption         =   "Forma de Pgto:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TamanhoAuto     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField2 
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   5220
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   370
         Caption         =   "Cód. Func.:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TamanhoAuto     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField41 
         Height          =   480
         Left            =   240
         TabIndex        =   1
         Top             =   2700
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   847
         Caption         =   "Descrição: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txthead 
         Height          =   390
         Left            =   9180
         TabIndex        =   2
         Top             =   1680
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   688
         Caption         =   "00000"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField47 
         Height          =   480
         Left            =   240
         TabIndex        =   3
         Top             =   3285
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   847
         Caption         =   "A importância supra de "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDescricao 
         Height          =   480
         Left            =   1770
         TabIndex        =   4
         Top             =   2700
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   847
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtValor 
         Height          =   765
         Left            =   3240
         TabIndex        =   5
         Top             =   3285
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1349
         Caption         =   ""
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtEntrada 
         Height          =   300
         Left            =   6540
         TabIndex        =   7
         Top             =   5400
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   529
         Caption         =   "ASSINATURA"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtAssinatura 
         Height          =   300
         Left            =   5100
         TabIndex        =   8
         Top             =   5100
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   529
         Formato         =   "dd/mm/yyyy"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   8
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtData 
         Height          =   300
         Left            =   7560
         TabIndex        =   9
         Top             =   2160
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TamanhoAuto     =   -1  'True
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf2 
         Height          =   300
         Left            =   3720
         TabIndex        =   10
         Top             =   540
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   529
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf1 
         Height          =   510
         Left            =   3720
         TabIndex        =   11
         Top             =   60
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   900
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf3 
         Height          =   300
         Left            =   3720
         TabIndex        =   12
         Top             =   780
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   529
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf4 
         Height          =   300
         Left            =   3720
         TabIndex        =   13
         Top             =   1020
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   529
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   630
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   1111
         Caption         =   "RECIBO DE SUPRIMENTOS"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField Ret 
         Height          =   5730
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   10107
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin VB.Image imgLogo 
         Height          =   1215
         Left            =   180
         Stretch         =   -1  'True
         Top             =   120
         Width           =   3315
      End
   End
End
Attribute VB_Name = "REL_ReciboSuprimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'arquivo .ini
Public cCfg As ConfigItem
Public oIni As Ini
Private Sub Form_Load()
On Error GoTo TrataErro

'abrindo arquivo .ini
'Set oIni = New Ini
'oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
'var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

'Dim Prt As Printer
'Dim oldPrinter As String

'Armazena o nome da impressora atual
'oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
'For Each Prt In Printers
'   If Prt.DeviceName = var_ImpNormal Then
'      Set Printer = Prt
'      Exit For
'   End If
'Next

'Relatorio.NomeImpressora = var_ImpNormal

Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)

rf1.Caption = r("fantasia")
rf2.Caption = r("razao")
rf3.Caption = r("endereco") & ", " & r("cidade") & "-" & r("estado")
rf4.Caption = "CNPJ: " & r("cnpj") & " - IE: " & r("ie") & " - TELEFONE: " & r("telefone")

If Not IsNull(r("caminho")) Then
If Dir$(r("caminho")) <> "" Then Set imgLogo.Picture = LoadPicture(r("caminho"))
End If

If r.State <> 0 Then r.Close
Set r = Nothing
Exit Sub
   
TrataErro:
End Sub


