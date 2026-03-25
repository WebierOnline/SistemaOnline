VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form Ordem_Servicos_Consulta_Geral_Servicos 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "SERVIÇOS"
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   Icon            =   "Ordem_Servicos_Consulta_Geral_Servicos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQuantPecas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   5400
      TabIndex        =   5
      Top             =   6600
      Width           =   675
   End
   Begin VB.TextBox txtTotalPecas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6120
      TabIndex        =   4
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox txtTotalServicos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6120
      TabIndex        =   3
      Top             =   3060
      Width           =   1575
   End
   Begin VB.TextBox txtQuantServicos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   5400
      TabIndex        =   2
      Top             =   3060
      Width           =   675
   End
   Begin VB.TextBox txtCodOS 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   555
      Left            =   6480
      TabIndex        =   0
      Top             =   7200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   979
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
      MICON           =   "Ordem_Servicos_Consulta_Geral_Servicos.frx":030A
      PICN            =   "Ordem_Servicos_Consulta_Geral_Servicos.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape4 
      Height          =   795
      Left            =   0
      Top             =   7080
      Width           =   7830
   End
   Begin VB.Shape Shape2 
      Height          =   3435
      Left            =   0
      Top             =   3540
      Width           =   7830
   End
   Begin VB.Shape Shape1 
      Height          =   3435
      Left            =   0
      Top             =   0
      Width           =   7830
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   0
      Top             =   3420
      Width           =   7840
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   0
      Top             =   6960
      Width           =   7845
   End
End
Attribute VB_Name = "Ordem_Servicos_Consulta_Geral_Servicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   Dim vTotal As Currency
   
   If txtCodOS.Text = "" Then Exit Sub
   
   'EXIBIR NO GRID - SERVIÇOS
   sSQL = "SELECT * FROM os_itens WHERE (cod_os = " & txtCodOS.Text & ") ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'MOSTRAR A QUANT DE REGISTROS - SERVIÇOS
   txtQuantServicos.Text = Format(totalRegistros, "00")
   
   'SOMAR OS REGISTROS - SERVIÇOS
   vTotal = 0
   sSQL = "SELECT IFNULL(SUM(total), 0) AS valor_total FROM os_itens WHERE (cod_os = " & txtCodOS.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then vTotal = r("valor_total")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   txtTotalServicos.Text = Format(vTotal, ocMONEY)
   
   'EXIBIR NO GRID - PEÇAS
   sSQL = "SELECT * FROM pedidos_itens WHERE (cod_os = " & txtCodOS.Text & ") ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'MOSTRAR A QUANT DE REGISTROS - PEÇAS
   txtQuantPecas.Text = Format(totalRegistros, "00")
   
   'SOMAR OS REGISTROS - PEÇAS
   vTotal = 0
   sSQL = "SELECT IFNULL(SUM(total), 0) AS valor_total FROM pedidos_itens WHERE (cod_os = " & txtCodOS.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then vTotal = r("valor_total")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   txtTotalPecas.Text = Format(vTotal, ocMONEY)
End Sub
