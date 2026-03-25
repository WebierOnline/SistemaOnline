VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form Estorno_ReabrirPedidos 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "PARCELAS DO PEDIDO"
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   10065
   Icon            =   "Estorno_ReabrirPedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3255
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5741
      _Version        =   393216
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   9825
      TabIndex        =   1
      Top             =   120
      Width           =   9855
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REABERTURA DO PEDIDO"
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
         Left            =   180
         TabIndex        =   2
         Top             =   120
         Width           =   4035
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   675
      Left            =   8580
      TabIndex        =   0
      Top             =   4140
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   1191
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
      MICON           =   "Estorno_ReabrirPedidos.frx":23D2
      PICN            =   "Estorno_ReabrirPedidos.frx":23EE
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
Attribute VB_Name = "Estorno_ReabrirPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Public Sub LoadInformacoes(ByVal Pedido As Long)
sSQL = "SELECT *, CASE cancelado WHEN 0 THEN '' ELSE 'SIM' END AS vCancelado , CASE status_pedido WHEN 0 THEN 'SIM' ELSE '' END AS vStatus " & _
   "FROM Pedidos_Reabertura WHERE (cod_pedido = " & Pedido & ") order by data, hora;"

'Debug.Print sSQL

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
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 1300
      .ColWidth(2) = 2000
      .ColWidth(3) = 1300
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      
      .TextMatrix(0, 1) = "COD_PEDIDO"
      .TextMatrix(0, 2) = "USUARIO"
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "DATA"
      .TextMatrix(0, 5) = "HORA"
      .TextMatrix(0, 6) = "CANCEL."
      .TextMatrix(0, 7) = "ABERTO"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      '.ColAlignment(2) = 1
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("COD_PEDIDO")
            .TextMatrix(.Rows - 1, 2) = rTabela("LOGIN")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("VLR_PEDIDO"), ocMONEY)
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("DATA"), "DD/MM/YY")
            .TextMatrix(.Rows - 1, 5) = Format(rTabela("HORA"), ocHORA)
            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("vCancelado"))
            .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("vStatus"))
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 6
         If .TextMatrix(i, 6) = "SIM" Then
            .CellForeColor = vbRed
         Else
            .CellForeColor = vbBlack
         End If
         .CellFontBold = True
      Next

      For i = 1 To .Rows - 1
         .Row = i
         .Col = 7
         If .TextMatrix(i, 7) = "SIM" Then
            .CellForeColor = vbRed
         Else
            .CellForeColor = vbBlack
         End If
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
   End With
End Sub


