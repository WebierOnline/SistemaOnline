VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_Aniversariantes 
   Caption         =   "Relatório de Aniversariantes"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "REL_Aniversariantes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   85.461
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.286
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      Top             =   3720
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   767
      Tipo            =   7
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A WOLCAR deseja a todos um feliz anivesário."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2295
         TabIndex        =   5
         Top             =   -60
         Width           =   7020
      End
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Top             =   3360
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   635
      Begin ReportX.ReportField rptAniversario 
         Height          =   345
         Left            =   9780
         TabIndex        =   4
         Top             =   -15
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         Campo           =   "VAR_NASC"
         Formato         =   "dd/mm"
         Caption         =   ""
         TipoCampo       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField4 
         Height          =   345
         Left            =   60
         TabIndex        =   7
         Top             =   0
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   609
         Campo           =   "NOME"
         Formato         =   "dd/mm"
         Caption         =   ""
         TipoCampo       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField6 
         Height          =   345
         Left            =   7320
         TabIndex        =   8
         Top             =   0
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   609
         Campo           =   "CELULAR"
         Formato         =   "dd/mm"
         Caption         =   ""
         TipoCampo       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
   End
   Begin ReportX.ReportSection rptDias 
      Align           =   1  'Align Top
      Height          =   3360
      Left            =   0
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   5927
      Tipo            =   2
      Begin ReportX.ReportField ReportField2 
         Height          =   225
         Left            =   45
         TabIndex        =   2
         Top             =   3105
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   397
         Caption         =   "NOME "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField5 
         Height          =   225
         Left            =   9780
         TabIndex        =   3
         Top             =   3105
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   397
         Caption         =   "ANIVERSÁRIO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   555
         Left            =   30
         TabIndex        =   1
         Top             =   2400
         Width           =   11520
         _ExtentX        =   20320
         _ExtentY        =   979
         Caption         =   "ANIVERSÁRIANTES DO MÊS"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   225
         Left            =   7320
         TabIndex        =   6
         Top             =   3120
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   397
         Caption         =   "CELULAR"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin VB.Image Image1 
         Height          =   2400
         Left            =   3960
         Picture         =   "REL_Aniversariantes.frx":030A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3585
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   60
      TabIndex        =   0
      Top             =   4260
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Regua           =   -1  'True
      Titulo          =   ""
      Registrado      =   0   'False
   End
End
Attribute VB_Name = "REL_Aniversariantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   'On Error GoTo TrataErro
   'Dim sSQL As String
   'Dim r As ADODB.Recordset
   
   'sSQL = "SELECT * FROM empresa ORDER BY fantasia LIMIT 0, 1;"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'rf1.Caption = r("fantasia")
   'rf2.Caption = r("razao")
   'rf3.Caption = r("endereco") & ", " & r("cidade") & "-" & r("estado")
   'rf4.Caption = "CNPJ: " & r("cnpj") & " - IE: " & r("ie") & " - TELEFONE: " & r("telefone")
   
   'If Not IsNull(r("caminho")) Then
   '   If Dir$(r("caminho")) <> "" Then Set imgLogo.Picture = LoadPicture(r("caminho"))
   'End If
   
   'If r.State <> 0 Then r.Close
   'Set r = Nothing
   'Exit Sub
   
TrataErro:
End Sub

Private Sub Rpx_MsgErro(Numero As Long)
   Dim Msg As String
   
   If Numero < 0 Then
      ' Mensagens de erro previstas
      Select Case Numero - vbObjectError
         Case 1001: Msg = "É necessário existir uma impressora instalada no Windows"
         Case 1002: Msg = "Não há registros a imprimir"
         Case 1003: Msg = "Não foi definida a seção de detalhe do relatório"
         Case 1004: Msg = "A configuração das seções de grupos está incorreta"
         Case 1005: Msg = "Foi definido um cursor do tipo Forward-Only para o recordset do relatório."
         Case 1006: Msg = "A página configurada para o relatório não possuí espaço suficiente para a impressão"
         Case 1007: Msg = "Já existe um relatório em andamento"
      End Select
      
      ShowMsg Msg, vbInformation
   Else
      ' Mensagens não previstas. Isso pode significar um erro
      ' interno no ReportX. Se isso acontecer, por favor reporte isso
      ' através de e-mail para ser corrigido.
      ShowMsg "Erro não previsto:" & Numero & vbCrLf & Error(Numero) & _
         IIf(Err.Number <> 0, vbCrLf + Err.Description, ""), vbCritical
   End If
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
   Rpx_MsgErro Numero
End Sub
