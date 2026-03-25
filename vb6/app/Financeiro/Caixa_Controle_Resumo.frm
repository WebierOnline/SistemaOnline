VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Caixa_Controle_Resumo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RESUMO DO CAIXA"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid_Prazo 
      Height          =   4905
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8652
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   6330
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8149
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
            TextSave        =   "14:19"
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
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      Height          =   675
      Left            =   60
      TabIndex        =   10
      Top             =   5040
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1191
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
      MICON           =   "Caixa_Controle_Resumo.frx":0000
      PICN            =   "Caixa_Controle_Resumo.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lbl6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10680
      TabIndex        =   14
      Top             =   6105
      Width           =   915
   End
   Begin VB.Label lbl06Tit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Left            =   10095
      TabIndex        =   13
      Top             =   6105
      Width           =   510
   End
   Begin VB.Label lbl5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10680
      TabIndex        =   12
      Top             =   5880
      Width           =   915
   End
   Begin VB.Label lbl05Tit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Left            =   10095
      TabIndex        =   11
      Top             =   5880
      Width           =   510
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10680
      TabIndex        =   9
      Top             =   4980
      Width           =   915
   End
   Begin VB.Label lbl01Tit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant.:"
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
      Left            =   9960
      TabIndex        =   8
      Top             =   4980
      Width           =   645
   End
   Begin VB.Label lbl03Tit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recebido:"
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
      Left            =   9720
      TabIndex        =   7
      Top             =   5430
      Width           =   885
   End
   Begin VB.Label lbl04Tit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Left            =   10095
      TabIndex        =   6
      Top             =   5655
      Width           =   510
   End
   Begin VB.Label lbl02Tit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
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
      Left            =   9825
      TabIndex        =   5
      Top             =   5205
      Width           =   780
   End
   Begin VB.Label lbl4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10680
      TabIndex        =   4
      Top             =   5655
      Width           =   915
   End
   Begin VB.Label lbl3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10680
      TabIndex        =   3
      Top             =   5430
      Width           =   915
   End
   Begin VB.Label lbl2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10680
      TabIndex        =   1
      Top             =   5205
      Width           =   915
   End
End
Attribute VB_Name = "Caixa_Controle_Resumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim printSQL As String
Dim sSQL As String
Dim r As ADODB.Recordset
Dim oCfg As ConfigItem
Dim vAluguelAtiva As Boolean
Dim vOSAtiva As Boolean


Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i  As Integer

If varTipoConsulta = "SUPRIMENTOS" Or varTipoConsulta = "SANGRIAS" Then
    With Grid_Prazo
       .Clear
       .Cols = 6
       .rows = 2
       
       .ColWidth(0) = 0
       .ColWidth(1) = 1600
       .ColWidth(2) = 4000
       .ColWidth(3) = 1000
       .ColWidth(4) = 900
       .ColWidth(5) = 1600

       ''SUPRIMENTO' AS ORIGEM, DESCRICAO, VALOR, DATA, FORMA_PGTO
       .TextMatrix(0, 1) = "ORIGEM"
       .TextMatrix(0, 2) = "DESCRICAO"
       .TextMatrix(0, 3) = "VALOR"
       .TextMatrix(0, 4) = "DATA"
       .TextMatrix(0, 5) = "FORMA PGTO."
       
       'colocar os cabeçalho em negrito
       For i = 0 To .Cols - 1
          .Col = i
          .Row = 0
          .CellFontBold = True
       Next
       
       .ColAlignment(1) = 3
       .Redraw = False
       i = 1
       
       If Not rTabela Is Nothing Then
          Do While Not rTabela.EOF
             .TextMatrix(.rows - 1, 1) = UCase(rTabela("ORIGEM"))
             .TextMatrix(.rows - 1, 2) = UCase(rTabela("DESCRICAO"))
             .TextMatrix(.rows - 1, 3) = Format(rTabela("VALOR"), ocMONEY)
             .TextMatrix(.rows - 1, 4) = Format(rTabela("DATA"), "dd/mm/yy")
             .TextMatrix(.rows - 1, 5) = rTabela("FORMA_PGTO")
             
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
          .Col = 3
          .CellForeColor = &H8000&
          .CellFontBold = True
       Next
       
       .rows = .rows - 1
       .Redraw = True
    End With
ElseIf varTipoConsulta = "RETIRADAS" Then
    With Grid_Prazo
       .Clear
       .Cols = 6
       .rows = 2
       
       .ColWidth(0) = 0
       .ColWidth(1) = 1600
       .ColWidth(2) = 4000
       .ColWidth(3) = 1000
       .ColWidth(4) = 900
       .ColWidth(5) = 900

       ''SUPRIMENTO' AS ORIGEM, DESCRICAO, VALOR, DATA, FORMA_PGTO
       .TextMatrix(0, 1) = "ORIGEM"
       .TextMatrix(0, 2) = "DESCRICAO"
       .TextMatrix(0, 3) = "VALOR"
       .TextMatrix(0, 4) = "DATA"
       .TextMatrix(0, 5) = "HORA"
       
       'colocar os cabeçalho em negrito
       For i = 0 To .Cols - 1
          .Col = i
          .Row = 0
          .CellFontBold = True
       Next
       
       .ColAlignment(1) = 3
       .Redraw = False
       i = 1
       
       If Not rTabela Is Nothing Then
          Do While Not rTabela.EOF
             .TextMatrix(.rows - 1, 1) = UCase(rTabela("ORIGEM"))
             .TextMatrix(.rows - 1, 2) = UCase(rTabela("DESCRICAO"))
             .TextMatrix(.rows - 1, 3) = Format(rTabela("VALOR"), ocMONEY)
             .TextMatrix(.rows - 1, 4) = Format(rTabela("DATA"), "dd/mm/yy")
             .TextMatrix(.rows - 1, 5) = Format(rTabela("HORA"), "HH:MM")
             
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
          .Col = 3
          .CellForeColor = &H8000&
          .CellFontBold = True
       Next
       
       .rows = .rows - 1
       .Redraw = True
    End With
Else
    With Grid_Prazo
       .Clear
       .Cols = 8
       .rows = 2
       
       .ColWidth(0) = 0
       .ColWidth(1) = 900
       .ColWidth(2) = 1000
       .ColWidth(3) = 4300
       .ColWidth(4) = 1000
       .ColWidth(5) = 900
       .ColWidth(6) = 1600
       .ColWidth(7) = 1000
       
       .TextMatrix(0, 1) = "PEDIDO"
       .TextMatrix(0, 2) = "TIPO"
       .TextMatrix(0, 3) = "CLIENTE"
       .TextMatrix(0, 4) = "VALOR"
       .TextMatrix(0, 5) = "DATA."
       .TextMatrix(0, 6) = "FORMA PGTO."
       .TextMatrix(0, 7) = "ORIGEM"
    
       
       'colocar os cabeçalho em negrito
       For i = 0 To .Cols - 1
          .Col = i
          .Row = 0
          .CellFontBold = True
       Next
       
       .ColAlignment(1) = 3
       .Redraw = False
       i = 1
       
       If Not rTabela Is Nothing Then
          Do While Not rTabela.EOF
             .TextMatrix(.rows - 1, 1) = Format(rTabela("varcodped"), "000000")
             .TextMatrix(.rows - 1, 2) = UCase(rTabela("ORIGEM"))
             .TextMatrix(.rows - 1, 3) = UCase(rTabela("varnomecliente"))
             .TextMatrix(.rows - 1, 4) = Format(rTabela("VARVALOR"), ocMONEY)
             .TextMatrix(.rows - 1, 5) = Format(rTabela("VARDATA"), "dd/mm/yy")
             
                If rTabela("FORMA_PGTO") <> "CARTAO" Then
                   .TextMatrix(.rows - 1, 6) = rTabela("FORMA_PGTO")
                Else
                    If rTabela("varTipocartao") = "D" Then
                        .TextMatrix(.rows - 1, 6) = rTabela("FORMA_PGTO") & " DÉBITO"
                    Else
                        .TextMatrix(.rows - 1, 6) = rTabela("FORMA_PGTO") & " CRÉDITO"
                    End If
    
                '.TextMatrix(.Rows - 1, 5) = rTabela("varFormaPgto") & " (" & rTabela("vartipocartao") & ")"
                   
                   
                End If
             
             
             '.TextMatrix(.Rows - 1, 6) = rTabela("FORMA_PGTO")
             .TextMatrix(.rows - 1, 7) = rTabela("varTIPO")
             
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
          .Col = 4
          .CellForeColor = &H8000&
          .CellFontBold = True
       Next
       
       .rows = .rows - 1
       .Redraw = True
    End With
End If

'somar e diferenciar
Dim soma As Currency
Dim QUANT As Integer
'Dim i As Integer

If varTipoConsulta = "CARTAO" Then
    'DEBITO
    soma = 0
    QUANT = 0
    With Grid_Prazo
       For i = 1 To .rows - 1
          If .TextMatrix(i, 6) = "CARTAO DÉBITO" And IsNumeric(.TextMatrix(i, 4)) Then
             soma = soma + CCur(.TextMatrix(i, 4))
             QUANT = QUANT + 1
          End If
       Next
    End With
    
    lbl01Tit.Caption = "Débito:"
    lbl02Tit.Caption = "Valor:"
    lbl1.Caption = Format(QUANT, "000")
    lbl2.Caption = Format(soma, "#,##0.00")
    
    'CREDITO
    soma = 0
    QUANT = 0
    With Grid_Prazo
       For i = 1 To .rows - 1
          If .TextMatrix(i, 6) = "CARTAO CRÉDITO" And IsNumeric(.TextMatrix(i, 4)) Then
             soma = soma + CCur(.TextMatrix(i, 4))
             QUANT = QUANT + 1
          End If
       Next
    End With
    
    lbl03Tit.Caption = "Crédito:"
    lbl04Tit.Caption = "Valor:"
    lbl3.Caption = Format(QUANT, "000")
    lbl4.Caption = Format(soma, "#,##0.00")
    lbl05Tit.Visible = False
    lbl5.Visible = False
ElseIf varTipoConsulta = "OUTROS" Then
    lbl01Tit.Caption = "Quant.:"
    lbl1.Caption = Format(Caixa_Controle_semOS.txtQuantAvulso.Text, "000")
    
    'TRANSFERENCIA
    soma = 0
    'QUANT = 0
    With Grid_Prazo
       For i = 1 To .rows - 1
          If .TextMatrix(i, 6) = "TRANSFERENCIA" And IsNumeric(.TextMatrix(i, 4)) Then
             soma = soma + CCur(.TextMatrix(i, 4))
             'QUANT = QUANT + 1
          End If
       Next
    End With
    
    lbl02Tit.Caption = "Transferência:"
    lbl2.Caption = Format(soma, "#,##0.00")
    
    'DEPOSITO
    soma = 0
    'QUANT = 0
    With Grid_Prazo
       For i = 1 To .rows - 1
          If .TextMatrix(i, 6) = "DEPOSITO" And IsNumeric(.TextMatrix(i, 4)) Then
             soma = soma + CCur(.TextMatrix(i, 4))
             'QUANT = QUANT + 1
          End If
       Next
    End With
    
    lbl03Tit.Caption = "Deposito:"
    lbl3.Caption = Format(soma, "#,##0.00")
    
    'BOLETO
    soma = 0
    'QUANT = 0
    With Grid_Prazo
       For i = 1 To .rows - 1
          If .TextMatrix(i, 6) = "BOLETO" And IsNumeric(.TextMatrix(i, 4)) Then
             soma = soma + CCur(.TextMatrix(i, 4))
             'QUANT = QUANT + 1
          End If
       Next
    End With
    
    lbl04Tit.Caption = "Boleto:"
    lbl4.Caption = Format(soma, "#,##0.00")
    
    'PIX
    soma = 0
    'QUANT = 0
    With Grid_Prazo
       For i = 1 To .rows - 1
          If .TextMatrix(i, 6) = "PIX" And IsNumeric(.TextMatrix(i, 4)) Then
             soma = soma + CCur(.TextMatrix(i, 4))
             'QUANT = QUANT + 1
          End If
       Next
    End With
    
    lbl05Tit.Visible = True
    lbl5.Visible = True
    lbl05Tit.Caption = "Pix:"
    lbl5.Caption = Format(soma, "#,##0.00")
    
ElseIf varTipoConsulta = "VENDAS" Then
    'DEBITO
    soma = 0
    QUANT = 0
    With Grid_Prazo
       For i = 1 To .rows - 1
          'If .TextMatrix(i, 6) = "CARTAO DÉBITO" And IsNumeric(.TextMatrix(i, 4)) Then
             soma = soma + CCur(.TextMatrix(i, 4))
             QUANT = QUANT + 1
          'End If
       Next
    End With
    
    lbl01Tit.Caption = "Quant.:"
    lbl02Tit.Caption = "Valor:"
    lbl1.Caption = Format(QUANT, "000")
    lbl2.Caption = Format(soma, "#,##0.00")
    lbl03Tit.Caption = ""
    lbl04Tit.Caption = ""
    lbl3.Caption = Format(0, "000")
    lbl4.Caption = Format(0, "#,##0.00")
    lbl05Tit.Visible = False
    lbl5.Visible = False
ElseIf varTipoConsulta = "PARCELAS" Then
    'DEBITO
    soma = 0
    QUANT = 0
    With Grid_Prazo
        For i = 1 To .rows - 1
            If .TextMatrix(i, 7) = "PARCELA" And IsNumeric(.TextMatrix(i, 4)) Then
               soma = soma + CCur(.TextMatrix(i, 4))
               QUANT = QUANT + 1
            End If
        Next
       
       If vAluguelAtiva = True Then
            Dim vSomaAluguel As Currency
            Dim vQuantAluguel As Integer
            vSomaAluguel = 0
            vQuantAluguel = 0
            For i = 1 To .rows - 1
              If .TextMatrix(i, 7) = "ALUGUEL" And IsNumeric(.TextMatrix(i, 4)) Then
                 vSomaAluguel = vSomaAluguel + CCur(.TextMatrix(i, 4))
                 vQuantAluguel = vQuantAluguel + 1
              End If
            Next
       End If

        If vOSAtiva = True Then
            Dim vSomaOS As Currency
            Dim vQuantOS As Integer
            vSomaOS = 0
            vQuantOS = 0

            For i = 1 To .rows - 1
              If .TextMatrix(i, 7) = "OS" And IsNumeric(.TextMatrix(i, 4)) Then
                 vSomaOS = vSomaOS + CCur(.TextMatrix(i, 4))
                 vQuantOS = vQuantOS + 1
              End If
            Next
        End If
       
       
    End With
    
    If vAluguelAtiva = False And vOSAtiva = False Then
        lbl01Tit.Caption = "Quant.:"
        lbl02Tit.Caption = "Valor:"
        lbl1.Caption = Format(QUANT, "000")
        lbl2.Caption = Format(soma, "#,##0.00")
        lbl03Tit.Caption = ""
        lbl04Tit.Caption = ""
        lbl3.Caption = Format(0, "000")
        lbl4.Caption = Format(0, "#,##0.00")
        lbl05Tit.Visible = False
        lbl5.Visible = False
    ElseIf vAluguelAtiva = True And vOSAtiva = False Then
        lbl01Tit.Caption = "Vendas:"
        lbl02Tit.Caption = "Valor:"
        lbl03Tit.Caption = "Aluguel"
        lbl04Tit.Caption = "Valor"
        lbl05Tit.Caption = ""
        lbl1.Caption = Format(QUANT, "000")
        lbl2.Caption = Format(soma, "#,##0.00")
        lbl3.Caption = Format(vQuantAluguel, "000")
        lbl4.Caption = Format(vSomaAluguel, "#,##0.00")
        lbl05Tit.Visible = False
        lbl5.Visible = False
    End If
ElseIf varTipoConsulta = "ALUGUEL" Then
    'DEBITO
    soma = 0
    QUANT = 0
    With Grid_Prazo
       For i = 1 To .rows - 1
          'If .TextMatrix(i, 6) = "CARTAO DÉBITO" And IsNumeric(.TextMatrix(i, 4)) Then
             soma = soma + CCur(.TextMatrix(i, 4))
             QUANT = QUANT + 1
          'End If
       Next
    End With
    
    lbl01Tit.Caption = "Quant.:"
    lbl02Tit.Caption = "Valor:"
    lbl1.Caption = Format(QUANT, "000")
    lbl2.Caption = Format(soma, "#,##0.00")
    lbl03Tit.Caption = ""
    lbl04Tit.Caption = ""
    lbl3.Caption = Format(0, "000")
    lbl4.Caption = Format(0, "#,##0.00")
    lbl05Tit.Visible = False
    lbl5.Visible = False
ElseIf varTipoConsulta = "HAVERES" Then
    'DEBITO
    soma = 0
    QUANT = 0
    With Grid_Prazo
       For i = 1 To .rows - 1
          'If .TextMatrix(i, 6) = "CARTAO DÉBITO" And IsNumeric(.TextMatrix(i, 4)) Then
             soma = soma + CCur(.TextMatrix(i, 4))
             QUANT = QUANT + 1
          'End If
       Next
    End With
    
    lbl01Tit.Caption = "Quant.:"
    lbl02Tit.Caption = "Valor:"
    lbl1.Caption = Format(QUANT, "000")
    lbl2.Caption = Format(soma, "#,##0.00")
    lbl03Tit.Caption = ""
    lbl04Tit.Caption = ""
    lbl3.Caption = Format(0, "000")
    lbl4.Caption = Format(0, "#,##0.00")
    lbl05Tit.Visible = False
    lbl5.Visible = False
ElseIf varTipoConsulta = "SUPRIMENTOS" Or varTipoConsulta = "SANGRIAS" Then
    'DEBITO
    soma = 0
    QUANT = 0
    With Grid_Prazo
       For i = 1 To .rows - 1
          'If .TextMatrix(i, 6) = "CARTAO DÉBITO" And IsNumeric(.TextMatrix(i, 4)) Then
             soma = soma + CCur(.TextMatrix(i, 3))
             QUANT = QUANT + 1
          'End If
       Next
    End With
    
    lbl01Tit.Caption = "Quant.:"
    lbl02Tit.Caption = "Valor:"
    lbl1.Caption = Format(QUANT, "000")
    lbl2.Caption = Format(soma, "#,##0.00")
    lbl03Tit.Caption = ""
    lbl04Tit.Caption = ""
    lbl3.Caption = Format(0, "000")
    lbl4.Caption = Format(0, "#,##0.00")
    lbl05Tit.Visible = False
    lbl5.Visible = False
ElseIf varTipoConsulta = "RETIRADAS" Then
    'DEBITO
    soma = 0
    QUANT = 0
    With Grid_Prazo
       For i = 1 To .rows - 1
          'If .TextMatrix(i, 6) = "CARTAO DÉBITO" And IsNumeric(.TextMatrix(i, 4)) Then
             soma = soma + CCur(.TextMatrix(i, 3))
             QUANT = QUANT + 1
          'End If
       Next
    End With
    
    lbl01Tit.Caption = "Quant.:"
    lbl02Tit.Caption = "Valor:"
    lbl1.Caption = Format(QUANT, "000")
    lbl2.Caption = Format(soma, "#,##0.00")
    lbl03Tit.Caption = ""
    lbl04Tit.Caption = ""
    lbl3.Caption = Format(0, "000")
    lbl4.Caption = Format(0, "#,##0.00")
    lbl05Tit.Visible = False
    lbl5.Visible = False
End If

'lblSubtotal.Caption = Format(SomaGrid(Grid_Prazo, 3), ocMONEY)
'lblEntrada.Caption = Format(SomaGrid(Grid_Prazo, 4), ocMONEY)
'lblTotal.Caption = Format(SomaGrid(Grid_Prazo, 7), ocMONEY)
End Sub
Private Sub FormatarGridPrazo(rTabela As ADODB.Recordset)
Dim i  As Integer

With Grid_Prazo
   .Clear
   .Cols = 9
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 900
   .ColWidth(2) = 4000
   .ColWidth(3) = 1100
   .ColWidth(4) = 1050
   .ColWidth(5) = 900
   .ColWidth(6) = 900
   .ColWidth(7) = 1000
   .ColWidth(8) = 1300
   
   .TextMatrix(0, 1) = "PEDIDO"
   .TextMatrix(0, 2) = "NOME DO CLIENTE"
   .TextMatrix(0, 3) = "SUBTOTAL"
   .TextMatrix(0, 4) = "RECEBIDO"
   .TextMatrix(0, 5) = "ACRESC."
   .TextMatrix(0, 6) = "DESC."
   .TextMatrix(0, 7) = "TOTAL"
   .TextMatrix(0, 8) = "TIPO"
   
   'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   .ColAlignment(1) = 3
   .Redraw = False
   i = 1
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = Format(rTabela("cod_pedido"), "000000")
         .TextMatrix(.rows - 1, 2) = UCase(rTabela("nome"))
         .TextMatrix(.rows - 1, 3) = Format(rTabela("SUBTOTAL"), ocMONEY)
         .TextMatrix(.rows - 1, 4) = Format(rTabela("varSomaRecebidos"), ocMONEY)
         .TextMatrix(.rows - 1, 5) = Format(rTabela("ValorAcrescReal"), ocMONEY)
         .TextMatrix(.rows - 1, 6) = Format(rTabela("ValorDescReal"), ocMONEY)
         .TextMatrix(.rows - 1, 7) = Format(rTabela("vartotal"), ocMONEY)
         .TextMatrix(.rows - 1, 8) = rTabela("pagamento")
         
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
      .Col = 3
      .CellForeColor = &H8000&
      .CellFontBold = True
   Next
   
   'Mudar as cores dependendo da situação
   For i = 1 To .rows - 1
      'For j = 0 To .Cols - 1
        .Row = i
        .Col = 4
         If .TextMatrix(i, 4) <> "0,00" Then
               .CellForeColor = vbRed
         Else
               .CellForeColor = vbBlack
        End If
        ' ElseIf .TextMatrix(i, 11) = "PAGO" Then
        '    .CellForeColor = vbBlue
        ' End If
      'Next
   Next
   
   
   .rows = .rows - 1
   .Redraw = True
End With

lbl02Tit.Caption = "Subtotal:"
lbl03Tit.Caption = "Recebido:"
lbl04Tit.Caption = "Total:"

lbl2.Caption = Format(SomaGrid(Grid_Prazo, 3), ocMONEY)
lbl3.Caption = Format(SomaGrid(Grid_Prazo, 4), ocMONEY)
lbl4.Caption = Format(SomaGrid(Grid_Prazo, 7) - SomaGrid(Grid_Prazo, 4), ocMONEY)
End Sub

Private Sub Mostrar_Outros()
Dim sSQL As String
Dim r As ADODB.Recordset
sSQL = "SELECT parcelas.COD_PEDIDO as varCodPed, 'PARCELA' AS ORIGEM, cliente.Nome as varNomeCliente, parcelas.VALOR_FINAL as varValor, parcelas.PAGAMENTO as varData, parcelas.FORMA_PGTO, parcelas.TIPO_CARTAO as varTipoCartao, parcelas.TIPO as varTipo " & _
    "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
    "WHERE (parcelas.FORMA_PGTO IN ('TRANSFERENCIA', 'DEPOSITO', 'BOLETO', 'FINANCEIRA', 'PIX')) AND (pedidos.TIPO_PEDIDO <> 'ORÇAMENTO') AND (pedidos.cancelado = 0) and (parcelas.codcaixa = " & varFluxoCodCaixa & ") AND parcelas.caixa = '" & varFluxoNomeCaixa & "' " & _
    "Union All " & _
    "SELECT parcelas.COD_PEDIDO as varCodPed, 'HAVER' AS ORIGEM, cliente.Nome as varNomeCliente, parcelas_haver.VALOR_HAVER as varValor, parcelas_haver.HAVER as varData, parcelas_haver.FORMA_PGTO, parcelas_haver.TIPO_CARTAO as varTipoCartao, parcelas_haver.TIPO as varTipo " & _
    "FROM parcelas_haver INNER JOIN parcelas AS parcelas ON parcelas_haver.COD_PARCELA = parcelas.CODIGO INNER JOIN pedidos AS pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente AS cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
    "WHERE (parcelas_haver.FORMA_PGTO IN ('TRANSFERENCIA', 'DEPOSITO', 'BOLETO', 'FINANCEIRA', 'PIX')) AND (pedidos.TIPO_PEDIDO <> 'ORÇAMENTO') AND (parcelas_haver.codcaixa = " & varFluxoCodCaixa & ") AND parcelas_haver.caixa = '" & varFluxoNomeCaixa & "' " & _
    "Union All " & _
    "SELECT '' as varCodPed, 'SUPRIMENTO' AS ORIGEM, DESCRICAO as varNomeCliente, VALOR as varValor, DATA as varData, FORMA_PGTO, '' as varTipoCartao, 'SUPRIMENTO' as varTipo " & _
    "FROM caixa_entrada " & _
    "WHERE (FORMA_PGTO IN ('TRANSFERENCIA', 'DEPOSITO', 'BOLETO', 'FINANCEIRA', 'PIX')) AND (codcaixa = " & varFluxoCodCaixa & ") AND caixa = '" & varFluxoNomeCaixa & "' "

Set r = dbData.OpenRecordset(sSQL)
'Debug.Print sSQL
printSQL = sSQL

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Mostrar_Vendas()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT parcelas.COD_PEDIDO as varcodped, 'PARCELA' AS ORIGEM, cliente.Nome as varNomecliente, parcelas.VALOR_FINAL as varValor, parcelas.PAGAMENTO as varData, parcelas.FORMA_PGTO, parcelas.TIPO_CARTAO as vartipocartao, parcelas.TIPO as vartipo " & _
    "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
    "WHERE (pedidos.TIPO_PEDIDO <> 'ORÇAMENTO') AND TIPO_PAGAMENTO <> 'À Prazo' and (pedidos.cancelado = 0) AND (parcelas.TIPO = 'VENDA') AND (parcelas.codcaixa = " & varFluxoCodCaixa & ") AND parcelas.caixa = '" & varFluxoNomeCaixa & "' "
Set r = dbData.OpenRecordset(sSQL)

printSQL = sSQL

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub Mostrar_Retiradas()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT 'SAÍDA' AS ORIGEM, DESCRICAO, VALOR, DATA, HORA " & _
    "FROM caixa_RETIRADA " & _
    "WHERE codcaixa = " & varFluxoCodCaixa & " AND caixa = '" & varFluxoNomeCaixa & "' "
Set r = dbData.OpenRecordset(sSQL)

printSQL = sSQL

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Mostrar_Sangrias()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT 'SANGRIAS' AS ORIGEM, DESCRICAO, VALOR, DATA, FONTE AS FORMA_PGTO " & _
    "FROM caixa_saida " & _
    "WHERE codcaixa = " & varFluxoCodCaixa & " AND caixa = '" & varFluxoNomeCaixa & "' AND FONTE = 'CAIXA ATUAL'"
Set r = dbData.OpenRecordset(sSQL)

printSQL = sSQL

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Mostrar_Suprimentos()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT 'SUPRIMENTO' AS ORIGEM, DESCRICAO, VALOR, DATA, FORMA_PGTO " & _
    "FROM caixa_entrada " & _
    "WHERE codcaixa = " & varFluxoCodCaixa & " AND caixa = '" & varFluxoNomeCaixa & "' "
Set r = dbData.OpenRecordset(sSQL)

printSQL = sSQL

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Mostrar_Haveres()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT parcelas.COD_PEDIDO as varcodped, 'HAVER' AS ORIGEM, cliente.Nome as varnomecliente, parcelas_haver.VALOR_HAVER as varValor, parcelas_haver.HAVER as varData, parcelas_haver.FORMA_PGTO, parcelas_haver.TIPO_CARTAO as vartipocartao, parcelas_haver.TIPO as vartipo " & _
    "FROM parcelas_haver INNER JOIN parcelas AS parcelas ON parcelas_haver.COD_PARCELA = parcelas.CODIGO INNER JOIN pedidos AS pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente AS cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
    "WHERE (pedidos.TIPO_PEDIDO <> 'ORÇAMENTO') AND (parcelas_haver.codcaixa = " & varFluxoCodCaixa & ") AND parcelas_haver.caixa = '" & varFluxoNomeCaixa & "' "
Set r = dbData.OpenRecordset(sSQL)

printSQL = sSQL

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub Mostrar_Servicos()
Dim sSQL As String
Dim r_Prazo As ADODB.Recordset

sSQL = "SELECT cliente.nome, pedidos.cod_pedido, (pedidos.TOTAL) as varTotal, pedidos.pagamento, pedidos.SUBTOTAL, pedidos.ValorDescReal, pedidos.ValorAcrescReal, pedidos.entrada, (SELECT ISNULL(SUM(VALOR_FINAL), 0) FROM parcelas WHERE (COD_PEDIDO = pedidos.COD_PEDIDO) AND (STATUS = 1)) AS varSomaRecebidos " & _
       "FROM pedidos INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
       "WHERE (pedidos.codcaixa = " & varFluxoCodCaixa & ") AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'OFICINA') and pedidos.cancelado = 0 AND pedidos.caixa = '" & varFluxoNomeCaixa & "' "
       'Debug.Print sSQL
       
Set r_Prazo = dbData.OpenRecordset(sSQL)

printSQL = sSQL

lbl01Tit.Caption = "Quant.:"
lbl1.Caption = Format(r_Prazo.RecordCount, "000")

FormatarGridPrazo r_Prazo

If r_Prazo.State <> 0 Then r_Prazo.Close
Set r_Prazo = Nothing
End Sub

Private Sub Mostrar_Aluguel()
Dim sSQL As String
Dim r_Prazo As ADODB.Recordset

sSQL = "SELECT cliente.nome, pedidos.cod_pedido, (pedidos.TOTAL) as varTotal, pedidos.pagamento, pedidos.SUBTOTAL, pedidos.ValorDescReal, pedidos.ValorAcrescReal, pedidos.entrada, (SELECT ISNULL(SUM(VALOR_FINAL), 0) FROM parcelas WHERE (COD_PEDIDO = pedidos.COD_PEDIDO) AND (STATUS = 1)) AS varSomaRecebidos " & _
       "FROM pedidos INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
       "WHERE (pedidos.codcaixa = " & varFluxoCodCaixa & ") AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'ALUGUEL') and pedidos.cancelado = 0 AND pedidos.caixa = '" & varFluxoNomeCaixa & "' "
       'Debug.Print sSQL
       
Set r_Prazo = dbData.OpenRecordset(sSQL)

printSQL = sSQL

lbl01Tit.Caption = "Quant.:"
lbl1.Caption = Format(r_Prazo.RecordCount, "000")

FormatarGridPrazo r_Prazo

If r_Prazo.State <> 0 Then r_Prazo.Close
Set r_Prazo = Nothing
End Sub

Private Sub Mostrar_Parcelas()
'Dim sSQL As String
'Dim r As ADODB.Recordset

sSQL = "SELECT parcelas.COD_PEDIDO as varcodped, 'PARCELA' AS ORIGEM, cliente.Nome as varnomecliente, parcelas.VALOR_FINAL as varValor, parcelas.PAGAMENTO as varData, parcelas.FORMA_PGTO, parcelas.TIPO_CARTAO as vartipocartao, parcelas.TIPO as vartipo " & _
    "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
    "WHERE (pedidos.TIPO_PEDIDO <> 'ORÇAMENTO') AND (pedidos.cancelado = 0) AND (parcelas.codcaixa = " & varFluxoCodCaixa & ") AND parcelas.caixa = '" & varFluxoNomeCaixa & "' AND (parcelas.TIPO IN ('PARCELA', 'ALUGUEL', 'OS')) AND (parcelas.STATUS = 1) "
Set r = dbData.OpenRecordset(sSQL)
Debug.Print sSQL
printSQL = sSQL

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Mostrar_Cartao()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT parcelas.COD_PEDIDO as varCodPed, 'PARCELA' AS ORIGEM, cliente.Nome as varNomeCliente, parcelas.VALOR_FINAL as varValor, parcelas.PAGAMENTO as varData, parcelas.FORMA_PGTO, parcelas.TIPO_CARTAO as varTipoCartao, parcelas.TIPO as varTipo " & _
    "FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
    "WHERE (pedidos.TIPO_PEDIDO <> 'ORÇAMENTO') AND (pedidos.cancelado = 0) AND (parcelas.FORMA_PGTO = 'CARTAO') AND (parcelas.codcaixa = " & varFluxoCodCaixa & ") AND parcelas.caixa = '" & varFluxoNomeCaixa & "' " & _
    "Union All " & _
    "SELECT parcelas.COD_PEDIDO as varCodPed, 'HAVER' AS ORIGEM, cliente.Nome as varNomeCliente, parcelas_haver.VALOR_HAVER as varValor, parcelas_haver.HAVER as varData, parcelas_haver.FORMA_PGTO, parcelas_haver.TIPO_CARTAO as varTipoCartao, parcelas_haver.TIPO as varTipo " & _
    "FROM parcelas_haver INNER JOIN parcelas AS parcelas ON parcelas_haver.COD_PARCELA = parcelas.CODIGO INNER JOIN pedidos AS pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO INNER JOIN cliente AS cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
    "WHERE (pedidos.TIPO_PEDIDO <> 'ORÇAMENTO') AND (pedidos.cancelado = 0) AND (parcelas_haver.FORMA_PGTO = 'CARTAO') AND (parcelas_haver.codcaixa = " & varFluxoCodCaixa & ") AND parcelas_haver.caixa = '" & varFluxoNomeCaixa & "'"
Set r = dbData.OpenRecordset(sSQL)

printSQL = sSQL

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Mostrar_APrazo()
Dim sSQL As String
Dim r_Prazo As ADODB.Recordset

sSQL = "SELECT cliente.nome , pedidos.cod_pedido, (pedidos.TOTAL) as varTotal, pedidos.pagamento, pedidos.SUBTOTAL, pedidos.ValorDescReal, pedidos.ValorAcrescReal, pedidos.entrada, (SELECT ISNULL(SUM(VALOR_FINAL), 0) FROM parcelas WHERE (COD_PEDIDO = pedidos.COD_PEDIDO) AND (STATUS = 1)) AS varSomaRecebidos " & _
       "FROM pedidos INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _
       "WHERE (pedidos.codcaixa = " & varFluxoCodCaixa & ") AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.TIPO_PEDIDO = 'VENDA') and pedidos.cancelado = 0 AND pedidos.caixa = '" & varFluxoNomeCaixa & "' "
       'Debug.Print sSQL
       
Set r_Prazo = dbData.OpenRecordset(sSQL)

printSQL = sSQL

lbl01Tit.Caption = "Quant.:"
lbl1.Caption = Format(r_Prazo.RecordCount, "000")

FormatarGridPrazo r_Prazo

If r_Prazo.State <> 0 Then r_Prazo.Close
Set r_Prazo = Nothing
End Sub

Private Sub cmdImprimir_Click()
Dim r As ADODB.Recordset
Dim vImpressoraNormal As String

'colocar o nome da maquina na barra de status
'Dim oIni As Ini

'Set oIni = New Ini
'oIni.Arquivo = appPathApp & "config.ini"
'vImpressoraNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
'Set oIni = Nothing

'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
vImpressoraNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = vImpressoraNormal Then
      Set Printer = Prt
      Exit For
   End If
Next

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Dim SomaItens As Currency

If varTipoConsulta = "PRAZO" Then
    Set REL_Caixa_Prazo.Relatorio.Recordset = r
    REL_Caixa_Prazo.lblTitulo.Caption = "RELATÓRIO DE CAIXA - VENDAS À PRAZO"
    REL_Caixa_Prazo.rfQuant.Caption = lbl1.Caption
    REL_Caixa_Prazo.rfSubTotal.Caption = Format(lbl2.Caption, ocMONEY)
    REL_Caixa_Prazo.rfEntrada.Caption = Format(lbl3.Caption, ocMONEY)
    REL_Caixa_Prazo.rfTotal.Caption = Format(lbl4.Caption, ocMONEY)
    
    REL_Caixa_Prazo.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
    REL_Caixa_Prazo.rfCodCaixa.Caption = varFluxoCodCaixa
    REL_Caixa_Prazo.rfNomeCaixa.Caption = varFluxoNomeCaixa
    
    REL_Caixa_Prazo.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Caixa_Prazo.Relatorio.Ativar
    Unload REL_Caixa_Prazo
ElseIf varTipoConsulta = "CARTAO" Then
    Set REL_Caixa_Venda.Relatorio.Recordset = r
    REL_Caixa_Venda.lblTitulo.Caption = "RELATÓRIO DE CAIXA - VENDAS NO CARTÃO"
    
    Dim SomaQuan As Integer
    SomaQuan = CInt(lbl1.Caption) + CInt(lbl3.Caption)
    REL_Caixa_Venda.rfQuant.Caption = SomaQuan
    
    SomaItens = CCur(lbl2.Caption) + CCur(lbl4.Caption)
    REL_Caixa_Venda.rfSubTotal.Caption = Format(SomaItens, ocMONEY)
    
    REL_Caixa_Venda.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
    REL_Caixa_Venda.rfCodCaixa.Caption = varFluxoCodCaixa
    REL_Caixa_Venda.rfNomeCaixa.Caption = varFluxoNomeCaixa
    
    'REL_Caixa_Prazo.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Caixa_Venda.Relatorio.Ativar
    Unload REL_Caixa_Venda

ElseIf varTipoConsulta = "OUTROS" Then
    Set REL_Caixa_Venda.Relatorio.Recordset = r
    REL_Caixa_Venda.lblTitulo.Caption = "RELATÓRIO DE CAIXA - VENDAS DE TRANSF./DEPOSITO/BOLETO/PIX"
    REL_Caixa_Venda.rfQuant.Caption = lbl1.Caption
    SomaItens = CCur(lbl2.Caption) + CCur(lbl3.Caption) + CCur(lbl4.Caption) + CCur(lbl5.Caption)
    
    REL_Caixa_Venda.rfSubTotal.Caption = Format(SomaItens, ocMONEY)
    
    REL_Caixa_Venda.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
    REL_Caixa_Venda.rfCodCaixa.Caption = varFluxoCodCaixa
    REL_Caixa_Venda.rfNomeCaixa.Caption = varFluxoNomeCaixa
    
    'REL_Caixa_Prazo.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Caixa_Venda.Relatorio.Ativar
    Unload REL_Caixa_Venda
ElseIf varTipoConsulta = "VENDAS" Then
    Set REL_Caixa_Venda.Relatorio.Recordset = r
    REL_Caixa_Venda.lblTitulo.Caption = "RELATÓRIO DE CAIXA - VENDAS"

    REL_Caixa_Venda.rfQuant.Caption = lbl1.Caption
    REL_Caixa_Venda.rfSubTotal.Caption = Format(lbl2.Caption, ocMONEY)
    
    REL_Caixa_Venda.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
    REL_Caixa_Venda.rfCodCaixa.Caption = varFluxoCodCaixa
    REL_Caixa_Venda.rfNomeCaixa.Caption = varFluxoNomeCaixa
    
    'REL_Caixa_Prazo.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Caixa_Venda.Relatorio.Ativar
    Unload REL_Caixa_Venda
ElseIf varTipoConsulta = "PARCELAS" Then
    Set REL_Caixa_Venda.Relatorio.Recordset = r
    REL_Caixa_Venda.lblTitulo.Caption = "RELATÓRIO DE CAIXA - PARCELAS"

    REL_Caixa_Venda.rfQuant.Caption = lbl1.Caption
    REL_Caixa_Venda.rfSubTotal.Caption = Format(lbl2.Caption, ocMONEY)
    
    REL_Caixa_Venda.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
    REL_Caixa_Venda.rfCodCaixa.Caption = varFluxoCodCaixa
    REL_Caixa_Venda.rfNomeCaixa.Caption = varFluxoNomeCaixa
    
    'REL_Caixa_Prazo.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Caixa_Venda.Relatorio.Ativar
    Unload REL_Caixa_Venda
ElseIf varTipoConsulta = "HAVERES" Then
    Set REL_Caixa_Venda.Relatorio.Recordset = r
    REL_Caixa_Venda.lblTitulo.Caption = "RELATÓRIO DE CAIXA - HAVERES"

    REL_Caixa_Venda.rfQuant.Caption = lbl1.Caption
    REL_Caixa_Venda.rfSubTotal.Caption = Format(lbl2.Caption, ocMONEY)
    
    REL_Caixa_Venda.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
    REL_Caixa_Venda.rfCodCaixa.Caption = varFluxoCodCaixa
    REL_Caixa_Venda.rfNomeCaixa.Caption = varFluxoNomeCaixa
    
    'REL_Caixa_Prazo.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Caixa_Venda.Relatorio.Ativar
    Unload REL_Caixa_Venda
ElseIf varTipoConsulta = "SUPRIMENTOS" Then
    Set REL_Caixa_SupriSangria.Relatorio.Recordset = r
    REL_Caixa_SupriSangria.lblTitulo.Caption = "RELATÓRIO DE CAIXA - SUPRIMENTOS"

    REL_Caixa_SupriSangria.rfQuant.Caption = lbl1.Caption
    REL_Caixa_SupriSangria.rfSubTotal.Caption = Format(lbl2.Caption, ocMONEY)
    
    REL_Caixa_SupriSangria.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
    REL_Caixa_SupriSangria.rfCodCaixa.Caption = varFluxoCodCaixa
    REL_Caixa_SupriSangria.rfNomeCaixa.Caption = varFluxoNomeCaixa
    
    'REL_Caixa_Prazo.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Caixa_SupriSangria.Relatorio.Ativar
    Unload REL_Caixa_SupriSangria
    
ElseIf varTipoConsulta = "SANGRIAS" Then
    Set REL_Caixa_SupriSangria.Relatorio.Recordset = r
    REL_Caixa_SupriSangria.lblTitulo.Caption = "RELATÓRIO DE CAIXA - SANGRIAS"

    REL_Caixa_SupriSangria.rfQuant.Caption = lbl1.Caption
    REL_Caixa_SupriSangria.rfSubTotal.Caption = Format(lbl2.Caption, ocMONEY)
    
    REL_Caixa_SupriSangria.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
    REL_Caixa_SupriSangria.rfCodCaixa.Caption = varFluxoCodCaixa
    REL_Caixa_SupriSangria.rfNomeCaixa.Caption = varFluxoNomeCaixa
    
    'REL_Caixa_Prazo.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Caixa_SupriSangria.Relatorio.Ativar
    Unload REL_Caixa_SupriSangria
ElseIf varTipoConsulta = "RETIRADAS" Then
    Set REL_Caixa_Retiradas.Relatorio.Recordset = r
    REL_Caixa_Retiradas.lblTitulo.Caption = "RELATÓRIO DE CAIXA - RETIRADAS"

    REL_Caixa_Retiradas.rfQuant.Caption = lbl1.Caption
    REL_Caixa_Retiradas.rfSubTotal.Caption = Format(lbl2.Caption, ocMONEY)
    
    REL_Caixa_Retiradas.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
    REL_Caixa_Retiradas.rfCodCaixa.Caption = varFluxoCodCaixa
    REL_Caixa_Retiradas.rfNomeCaixa.Caption = varFluxoNomeCaixa
    
    'REL_Caixa_Prazo.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Caixa_Retiradas.Relatorio.Ativar
    Unload REL_Caixa_Retiradas
ElseIf varTipoConsulta = "ALUGUEL" Then
    Set REL_Caixa_Aluguel.Relatorio.Recordset = r
    REL_Caixa_Aluguel.lblTitulo.Caption = "RELATÓRIO DE CAIXA - ALUGUEL"
    REL_Caixa_Aluguel.rfQuant.Caption = lbl1.Caption
    REL_Caixa_Aluguel.rfSubTotal.Caption = Format(lbl2.Caption, ocMONEY)
    REL_Caixa_Aluguel.rfEntrada.Caption = Format(lbl3.Caption, ocMONEY)
    REL_Caixa_Aluguel.rfTotal.Caption = Format(lbl4.Caption, ocMONEY)
    
    REL_Caixa_Aluguel.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
    REL_Caixa_Aluguel.rfCodCaixa.Caption = varFluxoCodCaixa
    REL_Caixa_Aluguel.rfNomeCaixa.Caption = varFluxoNomeCaixa
    
    REL_Caixa_Aluguel.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Caixa_Aluguel.Relatorio.Ativar
    Unload REL_Caixa_Aluguel
ElseIf varTipoConsulta = "SERVICOS" Then
    Set REL_Caixa_OS.Relatorio.Recordset = r
    REL_Caixa_OS.lblTitulo.Caption = "RELATÓRIO DE CAIXA - SERVICOS"
    REL_Caixa_OS.rfQuant.Caption = lbl1.Caption
    REL_Caixa_OS.rfSubTotal.Caption = Format(lbl2.Caption, ocMONEY)
    REL_Caixa_OS.rfEntrada.Caption = Format(lbl3.Caption, ocMONEY)
    REL_Caixa_OS.rfTotal.Caption = Format(lbl4.Caption, ocMONEY)
    
    REL_Caixa_OS.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
    REL_Caixa_OS.rfCodCaixa.Caption = varFluxoCodCaixa
    REL_Caixa_OS.rfNomeCaixa.Caption = varFluxoNomeCaixa
    
    REL_Caixa_OS.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Caixa_OS.Relatorio.Ativar
    Unload REL_Caixa_OS
End If

Me.Show
End Sub


Private Sub Form_Load()
Dim oIni As Ini

'mostrar os objetos de OS e/ou Aluguel
Dim oCfg As ConfigItem
'Dim bStatus As Boolean

'os
Set oCfg = sysConfig("os")    'Recupera a config deseja
'bStatus = CBool(oCfg.Value)   'Converte o valor para booleano
vOSAtiva = CBool(oCfg.Value)
Set oCfg = Nothing            'Destroi o objeto

'txtQuantDinheiroOS.Visible = bStatus 'Habilita/desabilida conforme valor
'txtTotalDinheiroOS.Visible = bStatus
'lblOS.Visible = bStatus

'aluguel
Set oCfg = sysConfig("aluguel")    'Recupera a config deseja
'bStatus = CBool(oCfg.Value)   'Converte o valor para booleano
vAluguelAtiva = CBool(oCfg.Value)
'Set oCfg = Nothing            'Destroi o objeto

'txtQuantDinheiroAluguel.Visible = bStatus 'Habilita/desabilida conforme valor
'txtTotalDinheiroAluguel.Visible = bStatus
'lblAluguel.Visible = bStatus



If varTipoConsulta = "PRAZO" Then
    Mostrar_APrazo
    lbl3.BackColor = &HC0C0FF
    
    lbl01Tit.Caption = "Quant.:"
    lbl02Tit.Caption = "Subtotal:"
    lbl03Tit.Caption = "Recebido:"
    lbl04Tit.Caption = "Total:"
    lbl05Tit.Caption = ""
    lbl06Tit.Caption = ""
    
    lbl5.Visible = False
    lbl05Tit.Visible = False
    lbl6.Visible = False
    lbl06Tit.Visible = False
    
    lbl01Tit.Left = 9960
    lbl02Tit.Left = 9825
    lbl03Tit.Left = 9720
    lbl04Tit.Left = 10095
    
    lbl01Tit.Top = 4980
    lbl02Tit.Top = 5205
    lbl03Tit.Top = 5430
    lbl04Tit.Top = 5655
    
    lbl1.Left = 10680
    lbl2.Left = 10680
    lbl3.Left = 10680
    lbl4.Left = 10680
    
    lbl1.Top = 4980
    lbl2.Top = 5205
    lbl3.Top = 5430
    lbl4.Top = 5655
    
    lbl1.Width = 915
    lbl2.Width = 915
    lbl3.Width = 915
    lbl4.Width = 915
    
    
ElseIf varTipoConsulta = "CARTAO" Then
    Mostrar_Cartao
    lbl3.BackColor = &H80000005
    lbl6.Visible = False
    lbl06Tit.Visible = False
ElseIf varTipoConsulta = "OUTROS" Then
    Mostrar_Outros
    lbl3.BackColor = &H80000005
    lbl6.Visible = False
    lbl06Tit.Visible = False
ElseIf varTipoConsulta = "VENDAS" Then
    Mostrar_Vendas
    lbl3.BackColor = &H80000005
    lbl3.Visible = False
    lbl03Tit.Visible = False
    lbl4.Visible = False
    lbl04Tit.Visible = False
    lbl5.Visible = False
    lbl05Tit.Visible = False
    lbl6.Visible = False
    lbl06Tit.Visible = False
ElseIf varTipoConsulta = "PARCELAS" Then
    Mostrar_Parcelas
    
    If vAluguelAtiva = False And vOSAtiva = False Then
        lbl3.Visible = False
        lbl4.Visible = False
        lbl5.Visible = False
        lbl03Tit.Visible = False
        lbl04Tit.Visible = False
        lbl05Tit.Visible = False
        lbl6.Visible = False
        lbl06Tit.Visible = False
        lbl3.BackColor = &H80000005
        lbl01Tit.Top = 5040
        lbl01Tit.Left = 10020
        lbl02Tit.Top = 5280
        lbl02Tit.Left = 9840
        lbl03Tit.Top = 5520
        lbl03Tit.Left = 9900
        lbl04Tit.Top = 5760
        lbl04Tit.Left = 10140
        lbl05Tit.Top = 6000
        lbl05Tit.Left = 10140
        'lbl06tit.Top = 5040
        'lbl06tit.Left = 10020
        
        lbl1.Top = 5040
        lbl1.Left = 10680
        lbl2.Top = 5280
        lbl2.Left = 10680
        lbl3.Top = 5520
        lbl3.Left = 10680
        lbl4.Top = 5760
        lbl4.Left = 10680
        lbl5.Top = 6000
        lbl5.Left = 10680
        'lbl6.Top = 5040
        'lbl6.Left = 10680
        lbl1.Width = 495
        lbl2.Width = 915
        lbl3.Width = 495
        lbl4.Width = 915
        lbl5.Width = 495
        lbl6.Width = 915
    ElseIf vAluguelAtiva = True And vOSAtiva = False Then
        lbl03Tit.Visible = False
        lbl04Tit.Visible = False
        lbl05Tit.Visible = False
        lbl06Tit.Visible = False
        
        lbl3.Visible = True
        lbl4.Visible = True
        lbl5.Visible = False
        lbl6.Visible = False
        
        lbl3.BackColor = &H80000005

        lbl01Tit.Caption = "Vendas:"
        lbl02Tit.Caption = "Aluguel:"
        lbl03Tit.Caption = "Ordem de Serviço:"
        
        lbl01Tit.Top = 5040
        lbl02Tit.Top = 5280
        lbl03Tit.Top = 5520
        
        lbl01Tit.Left = 9360
        lbl02Tit.Left = 9360
        lbl03Tit.Left = 8460
        
        lbl1.Top = 5040
        lbl2.Top = 5040
        lbl3.Top = 5280
        lbl4.Top = 5280
        'lbl5.Top = 5520
        'lbl6.Top = 5520
        
        'lbl1.Left = 9360
        'lbl2.Left = 9315
        'lbl3.Left = 8460
        
        lbl1.Left = 10080
        lbl2.Left = 10620
        lbl3.Left = 10080
        lbl4.Left = 10620
        'lbl5.Left = 10080
        'lbl6.Left = 10620
        
        lbl1.Width = 495
        lbl2.Width = 915
        lbl3.Width = 495
        lbl4.Width = 915
        lbl5.Width = 495
        lbl6.Width = 915
    End If
ElseIf varTipoConsulta = "HAVERES" Then
    Mostrar_Haveres
    lbl3.BackColor = &H80000005
    lbl3.Visible = False
    lbl03Tit.Visible = False
    lbl4.Visible = False
    lbl04Tit.Visible = False
    lbl5.Visible = False
    lbl05Tit.Visible = False
    lbl6.Visible = False
    lbl06Tit.Visible = False
ElseIf varTipoConsulta = "SUPRIMENTOS" Then
    Mostrar_Suprimentos
    lbl3.BackColor = &H80000005
    lbl3.Visible = False
    lbl03Tit.Visible = False
    lbl4.Visible = False
    lbl04Tit.Visible = False
    lbl5.Visible = False
    lbl05Tit.Visible = False
    lbl6.Visible = False
    lbl06Tit.Visible = False
ElseIf varTipoConsulta = "SANGRIAS" Then
    Mostrar_Sangrias
    lbl3.BackColor = &H80000005
    lbl3.Visible = False
    lbl03Tit.Visible = False
    lbl4.Visible = False
    lbl04Tit.Visible = False
    lbl5.Visible = False
    lbl05Tit.Visible = False
    lbl6.Visible = False
    lbl06Tit.Visible = False
ElseIf varTipoConsulta = "RETIRADAS" Then
    Mostrar_Retiradas
    lbl3.BackColor = &H80000005
    lbl3.Visible = False
    lbl03Tit.Visible = False
    lbl4.Visible = False
    lbl04Tit.Visible = False
    lbl5.Visible = False
    lbl05Tit.Visible = False
    lbl6.Visible = False
    lbl06Tit.Visible = False
ElseIf varTipoConsulta = "ALUGUEL" Then
    Mostrar_Aluguel
    lbl5.Visible = False
    lbl05Tit.Visible = False
    lbl6.Visible = False
    lbl06Tit.Visible = False
    lbl3.BackColor = &HC0C0FF
ElseIf varTipoConsulta = "SERVICOS" Then
    Mostrar_Servicos
    lbl5.Visible = False
    lbl05Tit.Visible = False
    lbl6.Visible = False
    lbl06Tit.Visible = False
    lbl3.BackColor = &HC0C0FF
End If
Caixa_Controle_Resumo.Caption = varTipoConsulta
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Caixa_Controle_Resumo.Hide
'Unload Me
'Caixa_Controle_semOS.Show
End Sub

Private Sub Grid_Prazo_DblClick()
If varTipoConsulta = "PRAZO" Then
    Parcelas_Consulta_Produtos.loadPedidos Grid_Prazo.TextMatrix(Grid_Prazo.Row, 1), Grid_Prazo.TextMatrix(Grid_Prazo.Row, 8)
ElseIf varTipoConsulta = "CARTAO" Then
    Parcelas_Consulta_Produtos.loadPedidos Grid_Prazo.TextMatrix(Grid_Prazo.Row, 1), Grid_Prazo.TextMatrix(Grid_Prazo.Row, 7)
ElseIf varTipoConsulta = "OUTROS" Then
    Parcelas_Consulta_Produtos.loadPedidos Grid_Prazo.TextMatrix(Grid_Prazo.Row, 1), Grid_Prazo.TextMatrix(Grid_Prazo.Row, 7)
ElseIf varTipoConsulta = "VENDAS" Then
    Parcelas_Consulta_Produtos.loadPedidos Grid_Prazo.TextMatrix(Grid_Prazo.Row, 1), Grid_Prazo.TextMatrix(Grid_Prazo.Row, 7)
ElseIf varTipoConsulta = "PARCELAS" Then
    Parcelas_Consulta_Produtos.loadPedidos Grid_Prazo.TextMatrix(Grid_Prazo.Row, 1), Grid_Prazo.TextMatrix(Grid_Prazo.Row, 7)
ElseIf varTipoConsulta = "HAVERES" Then
    Parcelas_Consulta_Produtos.loadPedidos Grid_Prazo.TextMatrix(Grid_Prazo.Row, 1), Grid_Prazo.TextMatrix(Grid_Prazo.Row, 7)
ElseIf varTipoConsulta = "SUPRIMENTOS" Then
    Mostrar_Suprimentos
ElseIf varTipoConsulta = "SANGRIAS" Then
    Mostrar_Sangrias
ElseIf varTipoConsulta = "ALUGUEL" Then
    Parcelas_Consulta_Produtos.loadPedidos Grid_Prazo.TextMatrix(Grid_Prazo.Row, 1), Grid_Prazo.TextMatrix(Grid_Prazo.Row, 8)
ElseIf varTipoConsulta = "SERVICOS" Then
    Parcelas_Consulta_Produtos.loadPedidos Grid_Prazo.TextMatrix(Grid_Prazo.Row, 1), Grid_Prazo.TextMatrix(Grid_Prazo.Row, 8)
End If

Parcelas_Consulta_Produtos.Show 1
End Sub


