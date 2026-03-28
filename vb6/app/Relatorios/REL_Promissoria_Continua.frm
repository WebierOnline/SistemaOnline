VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_Promissoria_Continua 
   Caption         =   "Promissórias"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   10.345
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   19.368
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   0
      TabIndex        =   12
      Top             =   5340
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   2
      Regua           =   -1  'True
      Escala          =   7
      Titulo          =   ""
      NomeImpressora  =   "IMPRESSORA1"
      Registrado      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   5272
      Left            =   0
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   9287
      Begin ReportX.ReportField txtNumero 
         Height          =   270
         Left            =   1440
         TabIndex        =   0
         Top             =   300
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtVencimento 
         Height          =   270
         Left            =   3180
         TabIndex        =   1
         Top             =   300
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtData 
         Height          =   270
         Left            =   330
         TabIndex        =   2
         Top             =   930
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtEmpresa 
         Height          =   270
         Left            =   2940
         TabIndex        =   3
         Top             =   1560
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   476
         Caption         =   " NOME DA EMPRESA AQUI"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCNPJ 
         Height          =   270
         Left            =   8280
         TabIndex        =   4
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
         Caption         =   "00.000.000/0000-00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtValor 
         Height          =   270
         Left            =   2460
         TabIndex        =   5
         Top             =   2160
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtLocal 
         Height          =   270
         Left            =   1770
         TabIndex        =   6
         Top             =   3585
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   476
         Caption         =   " URUÇUI - PI"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtEmissao 
         Height          =   270
         Left            =   8460
         TabIndex        =   7
         Top             =   4080
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtEmitente 
         Height          =   270
         Left            =   1800
         TabIndex        =   8
         Top             =   4020
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCPF 
         Height          =   270
         Left            =   1800
         TabIndex        =   9
         Top             =   4440
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtEndereco 
         Height          =   270
         Left            =   1800
         TabIndex        =   10
         Top             =   4800
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   476
         Caption         =   ""
         MostrarSeRepetir=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtParcela 
         Height          =   270
         Left            =   9240
         TabIndex        =   11
         Top             =   240
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
   End
End
Attribute VB_Name = "REL_Promissoria_Continua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCliente As ADODB.Recordset
Dim rsPedidos As ADODB.Recordset
Dim rsParcelas As ADODB.Recordset
Dim Cod_Pedido As Integer

Public Sub loadPromissoria(Pedido As Integer, PARCELA As Integer)
   Dim totalRegistros As Long
   
   'colocar o nome da maquina na barra de status
   Dim var_Impressora As String
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   Set oIni = Nothing
   
   Cod_Pedido = Pedido
   
   Set rsPedidos = dbData.OpenRecordset("SELECT * FROM pedidos WHERE (cod_pedido = " & Pedido & ");")
   Set rsCliente = dbData.OpenRecordset("SELECT * FROM cliente WHERE (codigo = " & rsPedidos("cod_cliente") & ");")
   
   If PARCELA = 0 Then
      Set rsParcelas = dbData.OpenRecordset("SELECT * FROM parcelas WHERE (cod_pedido = " & Pedido & ");", totalRegistros)
   Else
      Set rsParcelas = dbData.OpenRecordset("SELECT * FROM parcelas WHERE (cod_pedido = " & Pedido & ") AND (numero = " & PARCELA & ");", totalRegistros)
   End If
   
   Relatorio.NumeroRegistros = totalRegistros
   Relatorio.NomeImpressora = var_Impressora
   Relatorio.Ativar
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

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
   If Not rsParcelas.EOF Then rsParcelas.MoveNext
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)
   If Not rsParcelas.EOF Then
      txtNumero.Caption = Format(Cod_Pedido, "00000") & "/" & rsParcelas("numero")
      txtParcela.Caption = rsParcelas("valor")
      
      'Aqui é as data de vencimento
      txtData.Caption = LCase(NumeroExtenso(Format(rsParcelas("data"), "dd"), False)) & " de " & LCase(Format(rsParcelas("data"), "MMMM")) & " de " & Format(rsParcelas("data"), "yyyy")
      
      'Descomente aqui caso as data lá seja a data atual
      'txtData.Caption = UCase(numeroExtenso(Format(Date, "dd"), False)) & " de " & UCase(Format(Date, "MMMM")) & " de " & Format(Date, "yyyy")
      
      txtValor.Caption = NumeroExtenso(rsParcelas("valor"))
      
      'DADOS DO CLIENTE
      txtVencimento.Caption = Format(rsParcelas("data"), "dd/mm/yy")
      txtEmissao.Caption = Format(Date, "dd/mm/yyyy")
      txtEmitente.Caption = rsCliente("nome")
      txtCPF.Caption = IIf(IsNull(rsCliente("cpf")) = True, "", rsCliente("cpf"))
      txtEndereco.Caption = rsCliente("endereco") & " , " & rsCliente("bairro")
   End If
End Sub
