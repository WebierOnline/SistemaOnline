VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Ordem_Servicos_Consulta_Geral_Parcelas 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Consulta - Informações das Parcelas"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   6450
   Icon            =   "Ordem_Servicos_Consulta_Geral_Parcelas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3075
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   5424
      _Version        =   393216
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   555
      Left            =   5100
      TabIndex        =   1
      Top             =   3780
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
      MICON           =   "Ordem_Servicos_Consulta_Geral_Parcelas.frx":030A
      PICN            =   "Ordem_Servicos_Consulta_Geral_Parcelas.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   6435
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INFORMAÇÕES DAS PARCELAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   6240
   End
End
Attribute VB_Name = "Ordem_Servicos_Consulta_Geral_Parcelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub loadInformacoes(Pedido As Integer)
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT numero, data, valor, pagamento, CASE status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS pago " & _
      "FROM parcelas WHERE (cod_os = " & Pedido & ");"
   
   Set r = dbData.OpenRecordset(sSQL)
   FormatarGrid r
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 6
      .Rows = 2
          
      .ColWidth(0) = 0
      .ColWidth(1) = 1000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "No."
      .TextMatrix(0, 2) = "VENC."
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "PGTO"
      .TextMatrix(0, 5) = "STATUS"
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("numero")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("data"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("valor"), ocMONEY)
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("pagamento"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 5) = rTabela("pago")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 5
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub
