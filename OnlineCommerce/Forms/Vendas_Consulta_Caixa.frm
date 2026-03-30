VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Vendas_Consulta_Caixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONSULTA DE VENDA POR CAIXA"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "Vendas_Consulta_Caixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1395
      Left            =   60
      ScaleHeight     =   1335
      ScaleWidth      =   8115
      TabIndex        =   6
      Top             =   1020
      Width           =   8175
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ACRESC.:"
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
         Left            =   5580
         TabIndex        =   22
         Top             =   660
         Width           =   870
      End
      Begin VB.Label lblAcresc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5586
         TabIndex        =   21
         Top             =   900
         Width           =   1065
      End
      Begin VB.Label lblCodCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7140
         TabIndex        =   20
         Top             =   60
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6720
         TabIndex        =   18
         Top             =   900
         Width           =   1245
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Left            =   6720
         TabIndex        =   17
         Top             =   660
         Width           =   675
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4422
         TabIndex        =   16
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "DESC.:"
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
         Left            =   4440
         TabIndex        =   15
         Top             =   660
         Width           =   630
      End
      Begin VB.Label lblSubtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3195
         TabIndex        =   14
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Left            =   3180
         TabIndex        =   13
         Top             =   660
         Width           =   1110
      End
      Begin VB.Label lblTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "TIPO:"
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
         Left            =   1560
         TabIndex        =   11
         Top             =   660
         Width           =   510
      End
      Begin VB.Label lblForma 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   900
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "FORMA:"
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
         TabIndex        =   9
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   7770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CLIENTE:"
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
         TabIndex        =   7
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   4695
      Left            =   60
      ScaleHeight     =   4635
      ScaleWidth      =   8115
      TabIndex        =   4
      Top             =   2520
      Width           =   8175
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   4515
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   7964
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   8145
      TabIndex        =   0
      Top             =   60
      Width           =   8175
      Begin VB.TextBox txtCodPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   525
         Left            =   6540
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "000000"
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTA DE VENDA"
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
         Left            =   1500
         TabIndex        =   3
         Top             =   240
         Width           =   3390
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Vendas_Consulta_Caixa.frx":23D2
         Top             =   0
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   5940
         TabIndex        =   2
         Top             =   60
         Width           =   480
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   555
      Left            =   6600
      TabIndex        =   19
      Top             =   7320
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "&Sair"
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
      MICON           =   "Vendas_Consulta_Caixa.frx":8C18
      PICN            =   "Vendas_Consulta_Caixa.frx":8C34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   23
      Top             =   7995
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10319
            Text            =   "Online.Info - Informática"
            TextSave        =   "Online.Info - Informática"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "00:42"
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
End
Attribute VB_Name = "Vendas_Consulta_Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   '
End Sub

Private Sub lblCodCliente_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If lblCodCliente.Caption = "" Then
      lblCodCliente.Caption = "0"
      Exit Sub
   End If
   
   sSQL = "SELECT codigo, nome FROM cliente WHERE (codigo = " & lblCodCliente.Caption & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lblCliente.Caption = r("nome")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub txtCodPedido_Change()
   If txtCodPedido.Text = "" Then Exit Sub
   Mostrar_Dados
End Sub

Sub Mostrar_Dados()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT pedidos_itens.descricao, pedidos_itens.quantidade, pedidos_itens.preco, pedidos_itens.total AS var_totalitens, " & _
      "pedidos.total AS var_total, pedidos.tipo_desc AS var_tipodesc, pedidos.tipo_acrescimo AS var_tipoacresc, pedidos.* " & _
      "FROM pedidos INNER JOIN pedidos_itens ON pedidos.cod_pedido = pedidos_itens.cod_pedido " & _
      "WHERE (pedidos.cod_pedido = " & txtCodPedido.Text & ");"
   
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      lblCodCliente.Caption = r("cod_cliente")
      lblForma.Caption = UCase(r("tipo_pagamento"))
      lblTipo.Caption = r("pagamento")
      lblSubtotal.Caption = Format(r("subtotal"), ocMONEY)
      'lblDesc.Caption = RS!VALOR_DESC
      lblTotal.Caption = Format(r("var_total"), ocMONEY)
      
      If r("var_tipodesc") = "R" Then
         lblDesc.Caption = Format(r("valor_desc"), ocMONEY)
         lblAcresc.Caption = Format(r("valor_acrescimo"), ocMONEY)
      Else
         lblDesc.Caption = FormatNumber(r("valor_desc"), 2) & "%"
         lblAcresc.Caption = FormatNumber(r("valor_acrescimo"), 2) & "%"
      End If
   End If
   
   FormatarGrid r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'lbltotalGridProdutos.Caption = Format(SomaGrid(Grid, 9), "##,##0.00")
   'txtSubTotal.Text = Format(SomaGrid(Grid, 9), "##,##0.00")
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 4500
      .ColWidth(2) = 1050
      .ColWidth(3) = 1050
      .ColWidth(4) = 1050
      
      .TextMatrix(0, 1) = "DESCRICAO"
      .TextMatrix(0, 2) = "PREÇO"
      .TextMatrix(0, 3) = "QUANT."
      .TextMatrix(0, 4) = "TOTAL"
      .Redraw = False
      
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
            For i = 1 To .Rows - 1
               .Row = i
               .Col = 2:   .CellBackColor = &HC0FFFF
               .Col = 4:   .CellBackColor = &HC0C0FF
            Next
            
            .TextMatrix(.Rows - 1, 1) = rTabela("descricao")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("preco"), ocMONEY)
            .TextMatrix(.Rows - 1, 3) = rTabela("quantidade")
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("var_totalitens"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub
