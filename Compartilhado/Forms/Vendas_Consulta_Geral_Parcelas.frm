VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Vendas_Consulta_Geral_Parcelas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "PARCELAS DO PEDIDO"
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   9060
   Icon            =   "Vendas_Consulta_Geral_Parcelas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3015
      Left            =   60
      TabIndex        =   3
      Top             =   1080
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5318
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   8865
      TabIndex        =   1
      Top             =   120
      Width           =   8895
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Vendas_Consulta_Geral_Parcelas.frx":23D2
         Top             =   0
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARCELAS DO PEDIDO"
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
         TabIndex        =   2
         Top             =   300
         Width           =   3585
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   675
      Left            =   7620
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
      MICON           =   "Vendas_Consulta_Geral_Parcelas.frx":8C18
      PICN            =   "Vendas_Consulta_Geral_Parcelas.frx":8C34
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
Attribute VB_Name = "Vendas_Consulta_Geral_Parcelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub loadInformacoes(ByVal Pedido As Long)
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT NUMERO, DATA, PAGAMENTO, VALOR, JUROS, DESCONTO, VALOR_FINAL, FORMA_PGTO, CODCAIXA, CAIXA, " & _
        "CASE status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS varStatus,  " & _
        "(SELECT ISNULL(SUM(VALOR_HAVER), 0) FROM parcelas_haver WHERE (COD_PARCELA = parcelas.CODIGO)) AS varSomaHaveres " & _
        "FROM parcelas WHERE (cod_pedido = " & Pedido & ");"

Set r = dbData.OpenRecordset(sSQL)
'Debug.Print sSQL

FormatarGrid_Parcelas r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub FormatarGrid_Parcelas(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 11
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 450
      .ColWidth(2) = 900
      .ColWidth(3) = 1300
      .ColWidth(4) = 800
      .ColWidth(5) = 800
      .ColWidth(6) = 800
      .ColWidth(7) = 1100
      .ColWidth(8) = 900
      .ColWidth(9) = 900
      .ColWidth(10) = 1000
      
      'Valor , JUROS, DESCONTO, VALOR_FINAL
      .TextMatrix(0, 1) = "No."
      .TextMatrix(0, 2) = "VENC."
      .TextMatrix(0, 3) = "VLR BRUTO"
      .TextMatrix(0, 4) = "JUROS"
      .TextMatrix(0, 5) = "DESC."
      .TextMatrix(0, 6) = "HAVER"
      .TextMatrix(0, 7) = "VLR LIQ."
      .TextMatrix(0, 8) = "PGTO"
      .TextMatrix(0, 9) = "STATUS"
      .TextMatrix(0, 10) = "FORMA"
      
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
            .TextMatrix(.rows - 1, 1) = rTabela("numero")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("data"), "DD/MM/YY")
            .TextMatrix(.rows - 1, 3) = FormatNumber(rTabela("Valor"), 2)
            .TextMatrix(.rows - 1, 4) = FormatNumber(rTabela("JUROS"), 2)
            .TextMatrix(.rows - 1, 5) = FormatNumber(rTabela("DESCONTO"), 2)
            .TextMatrix(.rows - 1, 6) = FormatNumber(rTabela("varSomaHaveres"), 2)
            .TextMatrix(.rows - 1, 7) = FormatNumber(rTabela("VALOR_FINAL"), 2)
            .TextMatrix(.rows - 1, 8) = Format(rTabela("pagamento"), "DD/MM/YY")
            .TextMatrix(.rows - 1, 9) = rTabela("varStatus")
            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("FORMA_PGTO"))
            
            'If rTabela("pago") = "PAGO" Then
            '   .TextMatrix(.rows - 1, 3) = FormatNumber(rTabela("valor_final"), 2)
            'Else
            '   .TextMatrix(.rows - 1, 3) = Format(rTabela("valor"), ocMONEY)
            'End If
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 5
         If .TextMatrix(i, 5) = "PAGO" Then
            .CellForeColor = vbBlue
         Else
            .CellForeColor = vbRed
         End If
         
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
   End With
End Sub

