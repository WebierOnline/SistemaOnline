VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Vendas_Consulta_Geral_Pedidos 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "ITENS DO PEDIDO"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10650
   Icon            =   "Vendas_Consulta_Geral_Pedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3915
      Left            =   120
      TabIndex        =   10
      Top             =   1140
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   6906
      _Version        =   393216
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   120
      ScaleHeight     =   885
      ScaleWidth      =   10365
      TabIndex        =   8
      Top             =   180
      Width           =   10395
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ITENS DO PEDIDO"
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
         TabIndex        =   9
         Top             =   300
         Width           =   2745
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Vendas_Consulta_Geral_Pedidos.frx":23D2
         Top             =   0
         Width           =   1140
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   7320
      ScaleHeight     =   1185
      ScaleWidth      =   3165
      TabIndex        =   0
      Top             =   5100
      Width           =   3195
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   60
         Width           =   1755
      End
      Begin VB.Label lblTotalDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   420
         Width           =   1755
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUB-TOTAL:"
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
         TabIndex        =   4
         Top             =   60
         Width           =   1110
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCONTO:"
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
         TabIndex        =   3
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblTotalGeral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
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
         TabIndex        =   1
         Top             =   780
         Width           =   675
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   5100
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2143
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
      MICON           =   "Vendas_Consulta_Geral_Pedidos.frx":8C18
      PICN            =   "Vendas_Consulta_Geral_Pedidos.frx":8C34
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
Attribute VB_Name = "Vendas_Consulta_Geral_Pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub loadPedidos(ByVal Pedido As Long)
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim vSubTotal As Currency, vTotal As Currency
   
   sSQL = "SELECT descricao, quantidade, preco, valor_desc, (quantidade * preco) AS subtotal, " & _
      "CASE tipo_desc WHEN 'R' THEN (quantidade * preco) - valor_desc ELSE ((quantidade * preco) * (valor_desc / 100)) END AS total, codigo " & _
      "FROM pedidos_itens WHERE (cod_pedido = " & Pedido & ");"
   
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      vSubTotal = r("subtotal")
      vTotal = r("total")
   End If
   
   FormatarGrid_Itens r
   
   lblTotal.Caption = Format(vSubTotal, ocMONEY)
   lblTotalGeral.Caption = Format(vTotal, ocMONEY)
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & Pedido & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If r("tipo_desc") = "R" Then
      lblTotalDesc.Caption = Format(r("valor_desc"), ocMONEY)
   Else
       lblTotalDesc.Caption = FormatNumber(r("valor_desc"), 2) & "%"
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub FormatarGrid_Itens(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 7000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      
      .TextMatrix(0, 1) = "DESCRIÇÃO"
      .TextMatrix(0, 2) = "PREÇO"
      .TextMatrix(0, 3) = "QUANT."
      .TextMatrix(0, 4) = "TOTAL"
      
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
            .TextMatrix(.Rows - 1, 1) = rTabela("descricao")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("preco"), ocMONEY)
            .TextMatrix(.Rows - 1, 3) = rTabela("quantidade")
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("total"), ocMONEY)
         
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
   End With
End Sub
