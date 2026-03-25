VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormataCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

'tipo de dado do controle
Enum TipoDeCampo
    Character
    Numeric
    [Date ]
    Spinner
    Memo
    [Boolean]
    [Option ]
End Enum

'ATENCAO!!! O ".," DEVE ESTAR SEMPRE NAS PRIMEIRAS POSICOES....
'lista de caracteres que podem ser inseridos na máscara
Const C_INSERTAVEIS = ".,-+$/\()[]{}: "
Const C_NUMERICOS = "0123456789"                      'set de caracteres numéricos
Const TECLA_DEL = -7                                  'valor para a tecla DEL

Public vgRsLista As Object                        'DataControl da lst externa
Public vgValueAtual As Variant                    'valor atual da lst (do campo a ser capturado)
Public vgTemRS As Integer                         'flag - tem DataControl da lst externa
Public vgWndForm As Long                          'handle do form onde o controle está
Public vgNaoEClick As Integer                     'flag temporária
Public vgListaAberta As Boolean                   'lista (combo) está aberta

'propriedades da classe
Dim vgMask As String                              'máscara para o campo
Dim vgDataType As TipoDeCampo                     'tipo do campo
Dim vgStuffChar As String                         'caractere para justificar à esquerda
Dim vgSelectOnFocus As Boolean                    'seleciona o campo quando pegar o foco
Dim vgColorOnFocus As OLE_COLOR                   'cor quando o cursor estiver no campo
Dim vgColorNoFocus As OLE_COLOR                   'cor quando o cursor não estiver no campo
Dim vgFillStyleOnFocus As Integer                 'estilo do enchimento quando com foco
Dim vgFillStyleNoFocus As Integer                 'estilo do enchimento quando sem foco
Dim vgRowSource As String                         'nome do data control da lst externa
Dim vgListFields As String                        'nome dos campos para lst ou opções da lst interna
Dim vgBoundColumn As String                       'nome do campo para retornar da lista (value)
Dim vgFormatLostFocus As Boolean                  'flag - campo todo selecionado quando pegar o foco
Dim vgBookMark As Variant
Dim vgTag As String                               'tag do controle
Dim vgLocked As Boolean                           'se o controle esta bloqueado para edição
Dim vgCtPri As Control                            'controle principal
Dim vgCtScr As Control                            'scroll dos campos rotativos
Dim vgCtTit As Label                              'label do título dos campos
Dim vgCtShp As Shape                              'shape do formato que envolve o título
Dim vgCtFdo As Label                              'título do campo
Dim vgCtBot As CommandButton                      'botão extensor do campo
Dim vgFilter As String                            'filtro para lista externa
Dim vgAlignment As Long                           'alinhamento para o campo
Dim vgVisible As Boolean                          'visibilidade do campo
Dim vgValorOriginal As Variant                    'valor original do campo no RS
Dim vgSetValorOriginal As Boolean                 'se pode setar valor original do campo no RS

'dimensiona variaveis para a mascara
Dim vgMascara As String                           'máscaras
Dim vgMaiusculo As Integer                        'flag - maiúsculas?
Dim vgMinusculo As Integer                        'flag - minúsculas?
Dim vgSoAlfa As Integer                           'flag - só alfabético?
Dim vgQcDec As Integer                            'qde de casas decimais
Dim vgQcInt As Integer                            'qde de inteiros
Dim vgTemSpInt As Integer                         'se tem separadores de inteiros

'Dim vgfrmGMCale As New frmGMCale                  'classe ativa do momento
Dim vgListBoxSearch As String                     'string de procura na lista
Dim vgPriVez As Boolean                           'flag para evitar recursividade
Dim vgfrmCalcNaMem As Boolean                     'se o form de calculadora está na memória
Dim vgFezGotFocus As Boolean                      'flag ja fez o GotFocus

Dim FlagFocus As Boolean                          'flag para identificar o LostFocus quando abrindo a combo

'verifica se o botao para browse existe
Public Property Get BrowseVisible() As Integer
    If vgCtBot Is Nothing Then
        BrowseVisible = 0
    Else
        BrowseVisible = 1
    End If
End Property

'pega o nome do campo que relacionado com o banco
Public Property Get DataField() As String
    On Error Resume Next
    If Not vgCtPri Is Nothing Then
        DataField = vgCtPri.DataField
        If Err Then
            Err.Clear
            If Not vgCtFdo Is Nothing Then
                DataField = vgCtFdo.DataField
            End If
        End If
    End If
End Property

'filtro para lst externa - pega
Public Property Get Filter() As String
    Filter = vgFilter
End Property

'filtro para lst externa - coloca
Public Property Let Filter(ByVal New_Filter As String)
    vgFilter = New_Filter
End Property

'retorna alinhamento para o controle
Public Property Get Alignment() As Long
    Alignment = vgAlignment
End Property

'seta alinhamento para o controle
Public Property Let Alignment(ByVal New_Alignment As Long)
    On Error Resume Next
    If vgAlignment <> New_Alignment Then
        vgAlignment = New_Alignment
        vgCtPri.Alignment = vgAlignment
    End If
End Property

'abre recordset no no mesmo banco de dados da lista externa
Public Property Get OpenRecordSet(vgSQL As String) As Object
    RsOpen OpenRecordSet, vgSQL$
End Property

'muda plano do controle
Public Sub ZOrder(ByVal Index As Integer)
    On Error Resume Next                          'previne eventuais erros
    vgCtPri.ZOrder Index                          'tenta colocar ctl no plano indicado
End Sub

'foco no controle principal - coloca
Public Sub SetFocus()
    On Error Resume Next                          'previne eventuais erros
    vgCtPri.SetFocus                              'tenta colocar o foco no ctl principal
End Sub

'controle principal - coloca
Public Property Set CtPri(ByVal New_Ct As Control)
    Dim i As Boolean
    Set vgCtPri = New_Ct                          'controle principal
    vgVisible = True                              'visiblidade true
    On Error Resume Next                          'pode ser que o controle nao tenha esta propriedade
    i = New_Ct.Locked                             'esta bloqueado?
    If Err Then                                   'não tem a propriedade
        Err.Clear
    Else
        vgLocked = i                              'paga como esta
    End If
End Property

'controle principal - pega
Public Property Get CtPri() As Control
    Set CtPri = vgCtPri
End Property

'controle que será o título - coloca
Public Property Set CtTit(ByVal New_Ct As Label)
    Set vgCtTit = New_Ct
End Property

'controle resposável pelo formato do título - coloca
Public Property Set CtShp(ByVal New_Ct As Shape)
    Set vgCtShp = New_Ct
End Property

'cotrole que dará efeito 3d e será vinculado ao BD - coloca
Public Property Set CtFdo(ByVal New_Ct As Label)
    Set vgCtFdo = New_Ct
    vgCtFdo.ZOrder 0
End Property

'controle scroll de campos rotativos - coloca
Public Property Set CtScr(ByVal New_Ct As Control)
    Set vgCtScr = New_Ct
End Property

'controle botão extensor - coloca
Public Property Set CtBot(ByVal New_Ct As CommandButton)
    Set vgCtBot = New_Ct
End Property

'recordset que comanda a lst externa - pega
Public Property Get RecordSource() As Object
    If Not (vgRsLista Is Nothing) Then
        Set RecordSource = vgRsLista
    Else
        Set vgRsLista = New ADODB.Recordset
        Set RecordSource = vgRsLista
    End If
End Property

'recordset que comanda a lst externa - coloca
Public Property Set RecordSource(ByRef New_RecordSource As Object)
    If Not (vgRsLista Is Nothing) Then
        Set vgRsLista = Nothing
    End If
    Set vgRsLista = New_RecordSource
    If Not vgfrmGMCale.EncheLstRs(Me) Then
        Set vgRsLista = Nothing
    End If
    Set New_RecordSource = Nothing
End Property


'flag se o campo será reformatado após perder o foco - coloca
Public Property Let FormatLostFocus(ByVal New_FormatLostFocus As Boolean)
    vgFormatLostFocus = New_FormatLostFocus
End Property

'cor de fundo do controle - pega
Public Property Get BackColor() As OLE_COLOR
    BackColor = vgCtPri.BackColor
End Property

'cor de fundo do controle - coloca
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    vgCtPri.BackColor() = New_BackColor
End Property

'handle do controle - pega
Public Property Get hWnd() As Long
    On Error Resume Next                          'pode ser que o ctl principal não tenha handle
    hWnd = vgCtPri.hWnd                           'handle do ctl principal
    If Err Then                                   'ctl sem handle
        Err.Clear                                 'tira erro da memória
        hWnd = vgCtPri.Container.hWnd             'pega o handle do seu container
    End If
End Property

'marca para o controle normalmente usada em ctl do tipo optativo - pega
Public Property Get BookMark() As Variant
    BookMark = vgBookMark
End Property

'marca para o controle normalmente usada em ctl do tipo optativo - coloca
Public Property Let BookMark(ByVal New_BookMark As Variant)
    vgBookMark = New_BookMark
End Property

'largura do ctl - pega
Public Property Get Width() As Single
    If Not vgCtFdo Is Nothing Then
        Width = vgCtFdo.Width
    ElseIf Not vgCtPri Is Nothing Then
        Width = vgCtPri.Width
    End If
End Property

'largura do ctl - coloca
Public Property Let Width(ByVal New_Width As Single)
    vgCtPri.Width = New_Width
End Property

'Topo do ctl - pega
Public Property Get Top() As Single
    If Not vgCtFdo Is Nothing Then
        Top = vgCtFdo.Top
    ElseIf Not vgCtPri Is Nothing Then
        Top = vgCtPri.Top
    End If
End Property

'Topo do ctl - coloca
Public Property Let Top(ByVal New_Top As Single)
    Dim vgDif As Single
    If Not vgCtFdo Is Nothing Then
        vgDif = vgCtPri.Top - vgCtFdo.Top
        vgCtFdo.Top = New_Top
        vgCtPri.Top = New_Top + vgDif
    ElseIf Not vgCtPri Is Nothing Then
        vgCtPri.Top = New_Top
    End If
End Property

'Esquerdo do ctl - pega
Public Property Get Left() As Single
    If Not vgCtFdo Is Nothing Then
        Left = vgCtFdo.Left
    ElseIf Not vgCtPri Is Nothing Then
        Left = vgCtPri.Left
    End If
End Property

'Esquerdo do ctl - coloca
Public Property Let Left(ByVal New_Left As Single)
    Dim vgDif As Single
    If Not vgCtFdo Is Nothing Then
        vgDif = vgCtPri.Left - vgCtFdo.Left
        vgCtFdo.Left = New_Left
        vgCtPri.Left = New_Left + vgDif
    ElseIf Not vgCtPri Is Nothing Then
        vgCtPri.Left = New_Left
    End If
End Property

'altura do ctl - pega
Public Property Get Height() As Single
    Height = vgCtPri.Height
End Property

'altura do ctl - coloca
Public Property Let Height(ByVal New_Height As Single)
    vgCtPri.Height = New_Height
End Property

'caractere para ser mostrado no ctl quando digitando senha - pega
Public Property Get PasswordChar() As String
    PasswordChar = vgCtPri.PasswordChar
End Property

'caractere para ser mostrado no ctl quando digitando senha - coloca
Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    vgCtPri.PasswordChar() = New_PasswordChar
End Property

'cor do shape quando sem foco - pega
Public Property Get ColorNoFocus() As OLE_COLOR
    ColorNoFocus = vgColorNoFocus
End Property

'cor do shape quando sem foco - pega
Public Property Let ColorNoFocus(ByVal New_ColorNoFocus As OLE_COLOR)
    vgColorNoFocus = New_ColorNoFocus
    If Not (vgCtShp Is Nothing) Then              'ctl do formato já definido
        vgCtShp.FillColor() = vgColorNoFocus      'coloca cor nele
    End If
End Property

'cor do shape quando com foco - pega
Public Property Get ColorOnFocus() As OLE_COLOR
    ColorOnFocus = vgColorOnFocus
End Property

'cor do shape quando com foco - coloca
Public Property Let ColorOnFocus(ByVal New_ColorOnFocus As OLE_COLOR)
    vgColorOnFocus = New_ColorOnFocus
End Property

'controle travado (não editável) - pega
Public Property Get Locked() As Boolean
    If vgCtPri Is Nothing Then
        Locked = True
    Else
        Locked = (vgLocked Or vgCtPri.WhatsThisHelpID = -1)
    End If
End Property

'controle travado (não editável) - coloca
Public Property Let Locked(ByVal New_Locked As Boolean)
    vgLocked = New_Locked
    On Error Resume Next                          'pode ser que o ctl não tenha esta propriedade
    If Not vgCtPri Is Nothing Then
        vgCtPri.Locked = vgLocked                 'coloca ctl editável/não editável
    End If
End Property

'valor do click do botão do scroll - pega
Public Property Get SmallChange() As Integer
    If Not (vgCtScr Is Nothing) Then              'scroll já definido
        SmallChange = vgCtScr.SmallChange         'pega o valor de sua propriedade
    Else
        SmallChange = 0                           'senão, pega o default
    End If
End Property

'valor do click do botão do scroll - coloca
Public Property Let SmallChange(ByVal New_SmallChange As Integer)
    If Not (vgCtScr Is Nothing) Then              'scroll já definido
        vgCtScr.SmallChange() = New_SmallChange   'poe novo valor na sua propriedade
    End If
End Property

'valor quando click na barra de scroll - pega
Public Property Get LargeChange() As Integer
    If Not (vgCtScr Is Nothing) Then              'o ctl scroll já definido
        LargeChange = vgCtScr.LargeChange         'pega o valor de sua propriedade
    Else                                          'scroll não definido
        LargeChange = 0                           'valor = 0
    End If
End Property

'valor do click do scroll - coloca
Public Property Let LargeChange(ByVal New_LargeChange As Integer)
    If Not (vgCtScr Is Nothing) Then              'scroll já definido
        vgCtScr.LargeChange = New_LargeChange     'troca sua propriedade
    End If
End Property

Public Sub PosicionaLst(vgSoAbre As Boolean)
    Dim x As String, vgIndAtual As String, vgRegAtual As Variant, Ok As Boolean

    On Error Resume Next                          'previne um erro
    vgCtPri.Parent.AbreTabelaCbo vgCtPri.Index    'chama proc para abrir tabela
    If vgSoAbre Then Exit Sub

    On Error GoTo DeuErro
    If IsEmpty(vgValueAtual) Then
        GoSub PegaPriReg
    Else
        If vgRsLista.RecordCount > 0 Then
            If VarType(vgValueAtual) = 7 Then
                x$ = x$ + "CONVERT(VARCHAR, " + PoeColchetes$(vgBoundColumn) + ", 103) = "
            Else
                x$ = PoeColchetes$(vgBoundColumn) + " = "
            End If
            Select Case VarType(vgValueAtual)
                Case 8, 12                        'vbString e vbVariant
                    x$ = x$ + "'" + vgValueAtual + "'"
                Case 7
                    x$ = x$ + "'" + Format$(CDate(vgValueAtual), "dd/mm/yyyy") + "'"
                Case Else
                    x$ = x$ + CStr(vgValueAtual)
            End Select
            vgRsLista.MoveFirst
            vgRsLista.Find x$, 0, 1
            Ok = (vgRsLista.EOF = False And vgRsLista.BOF = False)
            If Ok Then
                On Error Resume Next
                vgfrmGMCale.grdLst.BookMark = vgRsLista.BookMark
                vgfrmGMCale.grdLst.ReBind
                If Err Then Err.Clear
            Else
                GoSub PegaPriReg
            End If
        End If
    End If
    Exit Sub

PegaPriReg:
    If vgRsLista.RecordCount > 0 Then             'vamos colocar no 1o.
        vgRsLista.MoveFirst                       'move para o primeiro reg
        If vgRsLista.EOF = False And vgRsLista.BOF = False Then 'registro localizado
            On Error Resume Next
            vgfrmGMCale.grdLst.BookMark = vgRsLista.BookMark    'sincronza grid com a tabela
            vgfrmGMCale.grdLst.ReBind
        Else
            vgfrmGMCale.Refresh
        End If
        If Err Then Err.Clear
    End If
    Return

DeuErro:
End Sub

'evento click no botao extensor
Public Sub botClick(Optional ByVal pTmPriCol As Integer)
    Dim p As POINTAPI, L As Long, T As Long, z As String, i As Long, h As Long, w As Long, _
        QI As Integer, k As String, Cancel As Boolean, Ct As Control, x As String, vgNVez As Integer
    vgCtPri.SetFocus
    'Define o tamanho da primeira coluna
    If pTmPriCol = 0 And DadosExtrasSepara(vgCtPri.Tag, "LstPriCol", 0) > 0 Then pTmPriCol = DadosExtrasSepara(vgCtPri.Tag, "LstPriCol", 0)

    vgNVez = 0
    Do While vgFezGotFocus = False And vgNVez < 100 'se ainda nao fez o gotfocus
        DoEvents                                    'vamos esperar um pouco
        vgNVez = vgNVez + 1
    Loop
    If vgFezGotFocus = False Then                   'esperamos e ainda não fez
        vgFezGotFocus = True                        'o gotfocus
        Exit Sub                                    'cai fora
    End If

    If vgLocked Or vgCtPri.WhatsThisHelpID = -1 Then
        Exit Sub                                    'se esta travado nao faz nada
    End If

    'base do controle
    If vgCtFdo Is Nothing Then                      'se não tem label de fundo (efeito 3d)
        Set Ct = vgCtPri                            'pega o próprio ctl primário
    Else                                            'senão,
        Set Ct = vgCtFdo                            'o label será a base
    End If
    p.x = Ct.Left / Screen.TwipsPerPixelX           'coordenadas da base
    p.y = Ct.Top / Screen.TwipsPerPixelY            'em relação a seu container
    ClientToScreen Ct.Container.hWnd, p             'coord em relação a tela

    'se for lst interna o botão irá abri uma lst abaixo do ctl base
    If Len(vgListFields) > 0 Then
        If ComLstExt Then
            PosicionaLst True                       'abre recordset se necessario
            i = (vgRsLista.RecordCount > 0)         'tem registro?
        Else
            i = (vgfrmGMCale.lst.ListCount > 0)     'lst interna tem itens?
        End If
        If i Then
            Set vgfrmGMCale.CtTransfer = Me         'ctl que chamou o calendário
            w = Ct.Width                            'tamanho do ctl base será tamanho da lst
            If ComLstExt Then
                Set vgfrmGMCale.Font = vgfrmGMCale.grdLst.Font 'fonte form = fonte da lst para cálculo de tamanho (textwidth e textheight)
                vgfrmGMCale.vgGrdCount = vgRsLista.RecordCount
                h = IIf(vgRsLista.RecordCount > 8, 8, vgRsLista.RecordCount)
                h = (h * (vgfrmGMCale.TextHeight("0") + Screen.TwipsPerPixelY)) + (2 * Screen.TwipsPerPixelY) 'altura maxima da lst
                vgfrmGMCale.lst.Visible = False
                vgfrmGMCale.grdLst.Visible = True
            Else
                Set vgfrmGMCale.Font = vgfrmGMCale.lst.Font   'fonte form = fonte da lst para cálculo de tamanho (textwidth e textheight)
                x$ = Substitui$(vgCtPri.Text, vgSubsTab$, vbTab, SO_UM)
                If vgfrmGMCale.lst.Text <> x$ Then            'se item selecionado
                    vgNaoEClick = -1                          'for diferente do ctl principal
                    vgfrmGMCale.lst.ListIndex = -1            'vamos procurá-lo na lista
                    For i = 0 To vgfrmGMCale.lst.ListCount - 1
                        If vgfrmGMCale.lst.List(i) = x$ Then  'achou...
                            vgfrmGMCale.lst.ListIndex = i     'seleciona
                            Exit For                          'e cai fora
                        End If
                    Next
                    vgNaoEClick = 0                           'flag se o evento click foi disparado
                End If
                h = IIf(vgfrmGMCale.lst.ListCount > 8, 8, vgfrmGMCale.lst.ListCount)
                h = (h * (vgfrmGMCale.TextHeight("0"))) + (2 * Screen.TwipsPerPixelY) 'altura maxima da lst
                vgfrmGMCale.lst.Visible = True
                vgfrmGMCale.grdLst.Visible = False
            End If
            L = (p.x * Screen.TwipsPerPixelX)                                         'coordenadas a onde a lst sera aberta
            T = (p.y * Screen.TwipsPerPixelY) + Ct.Height + Screen.TwipsPerPixelY
            If T + h > Screen.Height Then                                             'se lst nao cabe na tela abre para cima
                T = (p.y * Screen.TwipsPerPixelY) - h - Screen.TwipsPerPixelY
            End If
            vgfrmGMCale.Move L, T, w, h                                               'coloca lst na posição
            If ComLstExt Then                     'se for lst externa
                vgfrmGMCale.grdLst.ApproxCount = vgRsLista.RecordCount
                If Len(vgCtPri.Text) = 0 Then     'se nao tem item selecionado
                    If vgRsLista.RecordCount > 0 Then 'vamos colocar no 1o.
                        vgRsLista.MoveFirst           'move para o primeiro reg
                        If vgRsLista.EOF = False And vgRsLista.BOF = False Then 'registro localizado
                            On Error Resume Next
                            vgfrmGMCale.grdLst.ReBind
                            vgfrmGMCale.grdLst.BookMark = vgRsLista.BookMark    'sincronza grid com a tabela
                            If Err Then Err.Clear
                        End If
                    End If
                Else                              'caso contrario
                    vgfrmGMCale.grdLst.Refresh
                    PosicionaLst False            'posiciona no registro corrente
                End If
                On Error Resume Next              'cursor na coluna 0
                vgfrmGMCale.grdLst.Scroll -1 * vgfrmGMCale.grdLst.Col, 0
                vgfrmGMCale.AcertaTmCol pTmPriCol          'acerta tamanho da colunas
            End If
            DoEvents                              'deixa o windows trabalhar
            ChamaForm vgfrmGMCale                 'mostra a lst
            GotFocus                              'dispara evento GotFocus
        End If

    'campo data abre calendário
    ElseIf vgDataType = [Date ] Then
        Load vgfrmGMCale                          'le form para mostrar o calendário
        Set vgfrmGMCale.CtTransfer = Me           'ctl que chamou o calendário
        With vgfrmGMCale
            .Width = 1720                         'largura e altura do calendário
            .Height = 1545
            L = ((p.x * Screen.TwipsPerPixelX) + Ct.Width) - vgfrmGMCale.Width - 15
            If L < 0 Then L = 0                   'coordenadas do calendário
            T = (p.y * Screen.TwipsPerPixelY) + Ct.Height + 15
            If T + vgfrmGMCale.Height > Screen.Height Then
                T = (p.y * Screen.TwipsPerPixelY) - vgfrmGMCale.Width - 15
            End If
            .Left = L                             'posiciona o calendário
            .Top = T
            .IniciaCale MOSTRA_CALE               'monta calendário
            ChamaForm vgfrmGMCale                 'e mostra
            DoEvents                              'deixa o windows trabalhar
            GotFocus                              'dispara evento GotFocus
        End With

    'campo numéro abre calculadora
    ElseIf vgDataType = Numeric Or vgDataType = Spinner Then
        Load frmGMCalc                            'carrega calculadora na memória
        vgfrmCalcNaMem = True                     'flag, calculadora na memória
        With frmGMCalc                            'coloca no tag do form da calc algumas
            .Tag = vgCtPri.Text + "|"             'informações como: Valor inicial
            If vgQcDec = 0 Then                   'e máscara do campo
                .Tag = .Tag + String$(vgQcInt, "#")
            Else
                .Tag = .Tag + String$(vgQcInt - 1, "#") + "0." + String$(vgQcDec, "0")
            End If
            .IniciaCalc                           'prepara a calculadora
            L = ((p.x * Screen.TwipsPerPixelX) + Ct.Width) - frmGMCalc.Width - 15
            If L < 0 Then L = 0
            T = (p.y * Screen.TwipsPerPixelY) + Ct.Height + 15
            If T + frmGMCalc.Height > Screen.Height Then 'se não cabe embaixo coloca em cima
                T = (p.y * Screen.TwipsPerPixelY) - frmGMCalc.Width - 15
            End If
            .Left = L                                    'posiciona calculadora
            .Top = T
        End With
        ChamaForm frmGMCalc                              'mostra calculadora
        DoEvents                                         'deixa o windows trabalhar
        GotFocus                                         'dispara evento GotFocus
    End If
End Sub

'evento click no scroll
Public Sub ScrChange()
    Dim i As Integer, x As String
    If vgLocked Or vgCtPri.WhatsThisHelpID = -1 Then
        Exit Sub                                  'se esta travado nao faz nada
    End If
    On Error Resume Next                          'previne erros
    x$ = FormatText(Trim(Str(vgCtScr.Value)))     'valor do scroll
    vgCtPri.Text = x$                             'coloca no campo
    If Not vgPriVez Then vgCtPri.SetFocus         'coloca foco no clt primário
End Sub

'cor do texto - pega
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = vgCtPri.ForeColor
End Property

'cor do texto - coloca
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    vgCtPri.ForeColor() = New_ForeColor
End Property

'cor do título do campo - pega
Public Property Get CaptionColor() As OLE_COLOR
    If vgCtTit Is Nothing Then                    'se label do título não existe
        CaptionColor = 0                          'retorna preto como default
    Else
        CaptionColor = vgCtTit.ForeColor          'coloca cor do título
    End If
End Property

'cor do título do campo - coloca
Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
    If Not (vgCtTit Is Nothing) Then              'se título já existe
        vgCtTit.ForeColor() = New_CaptionColor    'coloca a cor no label do título
    End If
End Property

'Fonte do ctl principal - pega
Public Property Get Font() As Font
    Set Font = vgCtPri.Font
End Property

'Fonte do ctl principal - coloca
Public Property Set Font(ByVal New_Font As Font)
    Set vgCtPri.Font = New_Font
End Property

'Fonte do título - pega
Public Property Get CaptionFont() As Font
    If Not (vgCtTit Is Nothing) Then              'ctl do título já foi definido
        Set CaptionFont = vgCtTit.Font            'pega a sua fonte
    Else                                          'senão,
        Set CaptionFont = vgCtPri.Font            'pega fonte do ctl principal como default
    End If
End Property

'Fonte do título - Coloca
Public Property Set CaptionFont(ByVal New_CaptionFont As Font)
    If Not (vgCtTit Is Nothing) Then              'ctl do título já foi definido
        Set vgCtTit.Font = New_CaptionFont        'coloca a fonte nele
    End If
End Property

'verifica se o evento já foi disparado
Public Property Get PriVez() As Boolean
    PriVez = vgPriVez
End Property

'evento disparado quando o ctl principal muda o seu valor
Public Function Change() As Boolean
    Dim RetVal As Boolean, vgVl As Variant
    RetVal = False                                'evento chegou ate o fim?
    If Not vgPriVez Then                          'se não foi disparado de outro evento
        If vgDataType <> [Option ] And vgDataType <> [Boolean] Then
            If vgDataType = Spinner Then          'cp rotativo coloca o valor no scroll
                On Error Resume Next              'previne estouro no controle scroll
                vgCtScr.Value = ValBrasil(vgCtPri.Text)
                If Err Then                       'epa! Valor acima do permitido pelo scroll
                    If Val(vgCtPri.Text) > vgCtScr.Max Then 'vamos colocar o valor máximo
                        vgCtPri.Text = LTrim(Str(vgCtScr.Value)) 'do scroll no ctl principal
                    End If
                End If
            End If
            vgVl = Value                                         'valor atual
        End If
        RetVal = True                                            'evento chegou até o final...
    End If
    Change = RetVal                                              'retorna situação do evento
End Function

'evento disparado quando o foco sai do ctl principal
Public Sub LostFocus()
    Dim x As String, vgVl As Variant              'define variáveis locais
    On Error Resume Next                          'previne erros...
    If FlagFocus = False Then
        vgListaAberta = False
    End If
    FlagFocus = False

    'se o campo não for lógico/Optativo e flag de formatar campo após perder
    'o foco estiver ligada então vamos fazer...
    If vgDataType <> [Boolean] And vgDataType <> [Option ] And _
       vgDataType <> 99 And vgFormatLostFocus And _
       (vgCtPri.Parent.vgSituacao <> ACAO_NAVEGANDO Or _
       vgCtPri.Parent.vgTipo <> TP_TABELA) Then

        vgPriVez = True                           'evita recursividade de eventos
        x$ = FormatText(vgCtPri.Text)             'formata cp segundo a sua máscara
        vgPriVez = False
        If Err Then                               'Epa! um erro aconteceu...
            Err.Clear                             'limpa o erro da memória
            Class_Terminate                       'e termina esta classe
            Exit Sub                              'sai deste evento
        End If
        If x$ <> vgCtPri.Text Then                'cp formatado diferente do atual
            vgCtPri.Text = x$                     'coloca cp formatado no ctl principal
        End If
        vgPriVez = True                           'evita recursividade
        vgCtPri.Text = FormatText(vgCtPri.Text)   'reformata o texto no ctl principal
        vgPriVez = False
    End If
    If Not (vgCtShp Is Nothing) Then              'tem formato ao redor do título
        vgCtShp.FillColor = vgColorNoFocus        'troca a sua cor e
        vgCtShp.FillStyle = vgFillStyleNoFocus    'estilo quando não tem foco
    End If
    vgFezGotFocus = False                         'reseta flag
End Sub

'evento KeyDown
Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer, vgFezF4 As Boolean
    If KeyCode <> 0 Then                          'alguma tecla foi digitada
        If vgLocked Or vgCtPri.WhatsThisHelpID = -1 Then
            If KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight And _
               KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
               KeyCode <> vbKeyPageDown And KeyCode <> vbKeyPageUp And _
               KeyCode <> vbKeyEnd And KeyCode <> vbKeyHome Then
                KeyCode = 0
            End If
            Exit Sub
        End If

        'campo é uma lista (interna ou extena)
        i = (vgDataType <> [Boolean] And vgDataType <> [Option ])
        If i Then
            i = (Len(vgListFields) > 0 And vgCtPri.Enabled = True)
        End If
        If i Then                                 'Se o campo for lst
            Set vgfrmGMCale.CtTransfer = Me       'classe atual
            If KeyCode = vbKeyDown Then           'seta para baixo
                vgFezF4 = False                   'reseta flag F4
                vgListBoxSearch$ = ""             'limpa pesquisa por palavra
                If ComLstExt Then
                    If Not vgRsLista Is Nothing Then
                        PosicionaLst False        'posiciona no registro corrente
                        If vgRsLista.RecordCount > 0 Then
                            On Error Resume Next
                            If Len(vgfrmGMCale.grdLst.FirstRow) = 0 Then
                                vgfrmGMCale.grdLst.ReBind
                                vgfrmGMCale.SelecionaItem vgValueAtual
                            End If
                            If Len(Text) > 0 Then
                                vgfrmGMCale.grdLst.Row = vgfrmGMCale.grdLst.Row + 1
                            End If
                            If Err Then
                                Err.Clear
                                vgfrmGMCale.grdLst.Scroll 0, 1
                                vgfrmGMCale.grdLst.Row = vgfrmGMCale.grdLst.Row + 1
                            End If
                            If Err Then
                                Err.Clear
                            Else
                                vgRsLista.BookMark = vgfrmGMCale.grdLst.BookMark
                                vgfrmGMCale.PegaSelecao True
                                Change
                            End If
                        End If
                    End If
                Else
                    If vgfrmGMCale.lst.ListIndex < vgfrmGMCale.lst.ListCount - 1 Then
                        i = vgfrmGMCale.lst.ListIndex 'item atual
                        vgfrmGMCale.PoeItem i + 1     'seleciona o próximo
                        Change                        'executa evento change
                    End If
                End If
            ElseIf KeyCode = vbKeyUp Then             'seta para cima
                vgListBoxSearch$ = ""                 'limpa pesquisa por palavra
                If ComLstExt Then
                    If Not vgRsLista Is Nothing Then
                        PosicionaLst False            'posiciona no registro corrente
                        If vgRsLista.RecordCount > 0 Then
                            On Error Resume Next
                            If Len(vgfrmGMCale.grdLst.FirstRow) = 0 Then
                                vgfrmGMCale.grdLst.ReBind
                                vgfrmGMCale.SelecionaItem vgValueAtual
                            End If
                            vgfrmGMCale.grdLst.Row = vgfrmGMCale.grdLst.Row - 1
                            If Err Then
                                Err.Clear
                                vgfrmGMCale.grdLst.Scroll 0, -1
                                vgfrmGMCale.grdLst.Row = vgfrmGMCale.grdLst.Row - 1
                            End If
                            If Err Then
                                Err.Clear
                            Else
                                vgRsLista.BookMark = vgfrmGMCale.grdLst.BookMark
                                vgfrmGMCale.PegaSelecao True
                                Change
                            End If
                        End If
                    End If
                Else
                    If vgfrmGMCale.lst.ListIndex > 0 Then
                        i = vgfrmGMCale.lst.ListIndex 'item atual
                        vgfrmGMCale.PoeItem i - 1     'item anterior
                        Change                        'executa evento change
                    End If
                End If
            ElseIf KeyCode <= vbKeyHelp And KeyCode <> vbKeySpace And KeyCode <> vbKeyShift Then
                vgListBoxSearch$ = ""                 'limpa pesquisa por palavra
                If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then 'limpa lista
                    vgfrmGMCale.PoeItem -1                           'desmarca lista
                    Text = ""                                        'limpa ctl principal
                End If
            ElseIf KeyCode = vbKeyF4 Then                            'trata atalho F4
                vgFezF4 = True                                       'para abrir combo
            End If

            'so aceita estas teclas de controle
            If KeyCode <> vbKeyReturn And KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight And _
               KeyCode <> vbKeyEnd And KeyCode <> vbKeyHome Then
                KeyCode = 0
            End If
            If vgNaoEClick <> -2 Then
                Set vgfrmGMCale.CtTransfer = Nothing                 'libera memória da variável
            End If
        Else
            If KeyCode = vbKeyF4 Then                                'trata f4 com outros cps com extensor
                If Not vgCtBot Is Nothing Then                       'se tem extensor
                    vgFezF4 = True                                   'vamos abrir
                End If
            ElseIf KeyCode = vbKeyDelete Then                        'se pressionou DEL
                KeyCode = TrataTecla(TECLA_DEL, KeyCode)             'trata esta tecla...
            ElseIf KeyCode = vbKeyUp Then                            'seta cima,
                KeyCode = TrataTecla(-2, KeyCode)                    'vamos tratar adiante
            ElseIf KeyCode = vbKeyDown Then                          'seta para baixo
                KeyCode = TrataTecla(-1, KeyCode)                    'vamos tratar..
            ElseIf KeyCode = vbKeyTab And Shift = vbCtrlMask Then    'seta para baixo
                KeyCode = 0                                          'CTRL-TAB nao pode
            End If
        End If
        If vgFezF4 Then                                              'se f4,
            vgCtBot.Value = True                                     'força clique no extensor
        End If
    End If
End Sub

'evento GotFocus
Public Sub GotFocus()
    vgListBoxSearch$ = ""                         'começa pesquisa de palavra
    If Not (vgCtShp Is Nothing) Then              'tem formato em volta do título
        vgCtShp.FillColor = vgColorOnFocus        'coloca o cor e o tipo do shape
        vgCtShp.FillStyle = vgFillStyleOnFocus    'quando com foco
    End If

    'se não for campo memo, optativo e lógico e quer selecionar o campo
    If vgDataType <> Memo And vgDataType <> [Boolean] And vgDataType <> [Option ] And _
       vgDataType <> 99 And vgSelectOnFocus Then
        vgCtPri.SelStart = 0                      'cursor na posição 0
        vgCtPri.SelLength = Len(vgCtPri.Text)     'seleciona todo o texto do ctl
    End If
    On Error Resume Next                          'prepara para possivel erro...
    If mdiSISC.ActiveForm.vgCaracteristica = F_DADOS Then 'não é form de dados
        If Err Then
            Err.Clear
        Else                                              'se não esta incluindo e não pode alterar
            If vgCtPri.Parent.vgCaracteristica = F_DADOS Then
                If Err Then
                    Err.Clear
                Else
                    If mdiSISC.ActiveForm.vgSituacao = ACAO_NAVEGANDO And mdiSISC.ActiveForm.vgTemAlteracao = False Then
                        If vgCtPri.Tag <> "BRW" Then
                            vgCtPri.Locked = True         'bloqueia o campo
                        End If
                    Else                                  'se não
                        vgCtPri.Locked = Locked           'deixa como esta
                    End If
                End If
            End If
        End If
    End If
    mdiSISC.ActiveControl.TiraFoco                        'mas vamos tentar colocar
    vgFezGotFocus = True                                  'flag o evento chegou ao fim
End Sub

'evento KeyPress
Public Sub KeyPress(KeyAscii As Integer)
    If KeyAscii <> 0 Then                         'digitou algo...
        If vgLocked Or vgCtPri.WhatsThisHelpID = -1 Then
            KeyAscii = 0
            Exit Sub
        End If
        If Len(vgListFields) > 0 Then             'se for lst tenta achar na lst a tecla digitada
            If KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn And _
               KeyAscii <> vbKeyTab And KeyAscii <> 39 Then
                Set vgfrmGMCale.CtTransfer = Me   'classe atual

                'procura na lst
                vgListBoxSearch$ = vgfrmGMCale.SearchListBox(vgListBoxSearch$ + Chr$(KeyAscii), False)

                Set vgfrmGMCale.CtTransfer = Nothing 'tira referência da classe da memória
            End If
            If KeyAscii = vbKeyReturn Then           'teclou Enter
                SendK vbKeyTab                       'substitui por TAB
            End If
            KeyAscii = 0                             'estas teclas não fazem nada
        Else
            KeyAscii = TrataTecla(KeyAscii, KeyAscii) 'formata tecla digitada conforme a máscara
        End If
    End If
End Sub

'evento MoudeDown
Public Sub MouseDown()
    vgListBoxSearch$ = ""                         'recomeça pesquisa de palavras
End Sub

'mostra form de lst/calendário ou calculadora
Private Sub ChamaForm(F As Form)
    On Error Resume Next                          'previne erros...
    FlagFocus = True                              'flag para identificar o LostFocus quando estiver abrindo a lista
    vgWndForm = vgCtPri.Parent.hWnd               'handle do form onde o ctl está
    If Err = 0 Then                               'se tudo ok
        vgListaAberta = True                      'flag para identificar que a combo está aberta
        Set F.CtTransfer = Me                     'classe atual
        F.Enabled = True                          'habilita o form que iremos mostrar
        F.Show                                    'mostra o form
        If Err Then                               'erro - o form deve ser modal
            F.Show vbModal                        'vamos abrir também este de forma modal
        End If
    End If
End Sub

'flag se o texto será selecionado quando pegar foco - pega
Public Property Get SelectOnFocus() As Boolean
    SelectOnFocus = vgSelectOnFocus
End Property

'flag se o texto será selecionado quando pegar foco - coloca
Public Property Let SelectOnFocus(ByVal New_SelectOnFocus As Boolean)
    vgSelectOnFocus = New_SelectOnFocus
End Property

'valor mínimo do scroll - pega
Public Property Get Min() As Integer
    If Not (vgCtScr Is Nothing) Then              'scroll já definido
        Min = vgCtScr.Min                         'pega o seu valor
    Else                                          'caso contrário
        Min = 0                                   'pega o default
    End If
End Property

'valor mínimo do scroll - coloca
Public Property Let Min(ByVal New_Min As Integer)
    If Not (vgCtScr Is Nothing) Then              'scroll já definido
        On Error Resume Next
        vgCtScr.Min() = New_Min                   'atribui para a sua propriedade
    End If
End Property

'tamanho máximo de caracters no texto - pega
Public Property Get MaxLength() As Long
    On Error Resume Next
    MaxLength = vgCtPri.MaxLength
End Property

'tamanho máximo de caracters no texto - coloca
Public Property Let MaxLength(ByVal New_MaxLength As Long)
    If vgDataType <> Memo Then                    'campo memo não...
        vgCtPri.MaxLength() = New_MaxLength
    End If
End Property

'valor máximo para o scroll - pega
Public Property Get Max() As Integer
    If Not (vgCtScr Is Nothing) Then              'scroll já definido
        Max = vgCtScr.Max                         'pega o valor de sua propriedade
    Else
        Max = 32767                               'senão, pega o default (32767)
    End If
End Property

'valor máximo para o scroll - coloca
Public Property Let Max(ByVal New_Max As Integer)
    If Not (vgCtScr Is Nothing) Then              'scroll já definido
        On Error Resume Next
        vgCtScr.Max() = New_Max                   'coloca o valor na sua propriedade
    End If
End Property

'Método Refresh
Public Sub Refresh()
    If ComLstExt Then                             'cp tipo lst externa
        If Not vgRsLista Is Nothing Then
            Set vgfrmGMCale.CtTransfer = Me       'classe atual
            vgfrmGMCale.Remonta                   'tenta echer pegar o RS da lista
            Set vgfrmGMCale.CtTransfer = Nothing  'libera da memória referência da classe
        End If
    Else
        EncheRowSource                            'cp comun/lst interna
    End If
End Sub

'Método adiciona item na lst interna
Public Sub AddItem(ByVal New_Item As String)
    If Len(vgRowSource) = 0 Then                  'se não for lst externa
        If Len(vgListFields) > 0 Then             'se já tem algum item na lst
            vgListFields = vgListFields + "|"     'coloca o separador de itens
        End If
        vgListFields = vgListFields + New_Item    'coloca o novo item
        vgfrmGMCale.lst.AddItem New_Item          'e adiciona-o na lista
    End If
End Sub

'Método para limpar todos os itens da lst
Public Sub Clear()
    If Len(vgRowSource) = 0 Then                  'se não for lst externa
        If Len(vgListFields) > 0 Then             'se tem itens
            vgListFields = ""                     'tira
            vgfrmGMCale.lst.Clear                 'e mata os itens da lst
        End If
        vgCtPri.Text = ""                         'limpa o ctl principal
    End If
End Sub

'Método para remover um item da lista
Public Sub RemoveItem(ByVal New_Index As Integer)
    Dim i As Integer                              'dimensiona variáveis
    If Len(vgRowSource) = 0 Then                  'se não for lst externa
        vgfrmGMCale.lst.RemoveItem New_Index      'remove o itens da lst
        If vgfrmGMCale.lst.ListIndex = -1 Then    'se lst não selecionada
            vgCtPri.Text = ""                     'limpa contole principal
        End If
        vgListFields = ""                         'remonta variável da lst de itens
        For i = 0 To vgfrmGMCale.lst.ListCount - 1
            If Len(vgListFields) > 0 Then         'já tem um iten na lst
                vgListFields = vgListFields + "|" 'coloca um separador entre eles
            End If
            vgListFields = vgListFields + vgfrmGMCale.lst.List(i)
        Next
    End If
End Sub

'Item selecionado - pega
Public Property Get ListIndex() As Integer
    Dim x As String, i As Integer
    x$ = Value
    ListIndex = -1
    For i = 0 To vgfrmGMCale.lst.ListCount - 1
        If x$ = vgfrmGMCale.lst.List(i) Then
            ListIndex = i
            Exit For
        End If
    Next
End Property

'Item selecionado - coloca
Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    If vgfrmGMCale.OqMostrar = MOSTRA_NAO Then    'se ainda não encheu a lst
        EncheRowSource                            'vamos fazer agora...
    End If
    Set vgfrmGMCale.CtTransfer = Me               'classe atual
    vgfrmGMCale.PoeItem New_ListIndex             'coloca item requerido
    Set vgfrmGMCale.CtTransfer = Nothing          'libera da memória referência da classe
End Property

'quantidade de itens da lst - pega
Public Property Get ListCount() As Integer
    ListCount = vgfrmGMCale.lst.ListCount
End Property

'texto do item da lista - pega
Public Property Get List(ByVal Index As Integer) As String
    List = vgfrmGMCale.lst.List(Index)
End Property

'texto do item da lista - coloca
Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    vgfrmGMCale.lst.List(Index) = New_List

    'se for o item selecionado coloca no ctl principal também
    If vgfrmGMCale.lst.ListIndex = Index Then
        vgCtPri.Text = Substitui$(New_List, vbTab, vgSubsTab$, SO_UM)
    End If
End Property

'variável de usuário - pega
Public Property Get Tag() As String
    Tag = vgTag
End Property

'variável de usuário - coloca
Public Property Let Tag(ByVal New_Tag As String)
    vgTag = New_Tag
End Property

'texto selecionado no ctl principal - pega
Public Property Get SelText() As String
    SelText = vgCtPri.SelText
End Property

'texto selecionado no ctl principal - coloca
Public Property Let SelText(ByVal New_SelText As String)
    vgCtPri.SelText() = New_SelText
End Property

'Posição do cursor no ctl principal - pega
Public Property Get SelStart() As Long
    SelStart = vgCtPri.SelStart
End Property

'Posiçao do cursor no ctl principal - coloca
Public Property Let SelStart(ByVal New_SelStart As Long)
    vgCtPri.SelStart() = New_SelStart
End Property

'quantidade de caractere para seleção - pega
Public Property Get SelLength() As Long
    SelLength = vgCtPri.SelLength
End Property

'quantidade de caractere para seleção - coloca
Public Property Let SelLength(ByVal New_SelLength As Long)
    vgCtPri.SelLength() = New_SelLength
End Property

'texto do ctl principal - pega
Public Property Get Text() As String

    'campo lógico ou optativo não tem textbox
    If vgDataType = [Option ] Then
        Text = vgCtPri.Caption                    'pega o caption do ctl principal
    Else
        Text = vgCtPri.Text                       'outros tipos de campos pega da textbox (ctl principal)
    End If
End Property

'texto do ctl principal - coloca
Public Property Let Text(ByVal New_Text As String)
    Dim x As String
    If New_Text = "" Then                         'texto é tudo branco
        x$ = ""
    Else                                          'senão,
        x$ = FormatText(New_Text)                 'formata segundo sua máscara
    End If

    'nos campos lógicos ou optativos...
    If vgDataType = [Option ] Then
        vgCtPri.Caption = x$                      'usa a propriedade caption
    Else
        If ComLstExt And Len(x$) = 0 Then         'se for lst externa
            vgValueAtual = Empty                  'tira valor
        End If
        vgCtPri.Text = x$                         'senão, usa a propriedade texto
    End If
End Property

'valor do controle - pega
Public Property Get Value() As Variant

    'lst externa
    If ComLstExt Then
        If IsEmpty(vgValueAtual) Then             'se não tem item selecionado
            Select Case vgDataType                'de acordo com o seu tipo vamos
                Case Character, Memo              'retornar o seu valor
                    Value = ""                    'caractere ou memo: em branco
                Case Numeric, Spinner, [Option ]
                    Value = 0                     'numérico, rotativo ou optativo: 0
                Case [Boolean]
                    Value = False                 'lógico: falso
                Case [Date ]
                    Value = Null                  'data: nulo
            End Select
        Else                                      'tem item selecionado
            Value = vgValueAtual                  'pega o valor do campo de retorno (BoundColumn)
        End If
    Else                                          'não é lst externa
        Select Case vgDataType
            Case Character, Memo                  'caractere ou memo
                Value = vgCtPri.Text              'retorna o que esta no ctl principal
            Case Numeric, Spinner                 'numérico ou rotativo
                If Len(vgListFields) > 0 And vgDataType = Numeric Then 'lst interna
                    Value = vgfrmGMCale.lst.ListIndex                  'retorna número do item selecionado
                Else                              'caso contrário
                    Value = CtoN(vgCtPri.Text())  'transforma o texto do ctl principal em número
                End If
            Case [Boolean]                        'lógico
                Value = vgCtPri.Value             'o valor que está checado
            Case [Option ]                        'optativo
                Value = vgCtPri.Value             'o valor que está marcado
            Case [Date ]                          'data
                If Len(vgCtPri.Text()) = 0 Or IsDate(vgCtPri.Text()) = False Then
                    Value = Null                  'data inválido ou em branco
                Else

                    'se a data digitada no texto pode ser transformada para data
                    'vamos retorná-la, senão, retorna nula (data em branco)
                    If IsDate(Format$(vgCtPri.Text(), TrataMaskData$(vgMask))) Then
                        Value = CDate(Format$(vgCtPri.Text(), TrataMaskData$(vgMask)))
                    Else
                        Value = Null
                    End If
                End If
        End Select
    End If
End Property

'valor original do campo - pega
Public Property Get OriginalValue() As Variant
    OriginalValue = vgValorOriginal
End Property

'se pode setar valor original do campo - pega
Public Property Get SetOriginalValue() As Boolean
    SetOriginalValue = vgSetValorOriginal
End Property

'se pode setar valor original do campo - coloca
Public Property Let SetOriginalValue(ByVal NewValue As Boolean)
    vgSetValorOriginal = NewValue
End Property

'valor do controle - coloca
Public Property Let Value(ByVal New_Value As Variant)
    Dim x As String, vgPV As Boolean, vgVl As Variant  'cria variávies na memória
    vgPV = vgPriVez                                    'situação atual que evita recursividade
    vgPriVez = True                                    'força flag para não ter recursividade
    On Error GoTo DeuErro
    If ComLstExt Then                                  'é lst externa
        vgValueAtual = New_Value                       'valor do campo de retorno (BoundColumn)
        Set vgfrmGMCale.CtTransfer = Me                'classe atual
        vgCtPri.Text() = vgfrmGMCale.SelecionaItem(vgValueAtual) 'seleciona na lst
        Set vgfrmGMCale.CtTransfer = Nothing                     'libera memória
    ElseIf vgDataType = Spinner Then                             'campo rotativo
        If Not (vgCtScr Is Nothing) Then                         'tem scroll associado
            vgCtScr.Value() = Val(CStr(New_Value))               'coloca o valor no scroll
            vgCtPri.Text() = vgCtScr.Value()                     'coloca o valor no controle
        End If
    ElseIf vgDataType = [Date ] Then                             'campo data coloca formatando
        vgCtPri.Text() = Format$(Trim(New_Value), TrataMaskData$(vgMask))
    ElseIf vgDataType = [Boolean] Then                           'campo lógico
        If vgCtPri.Value <> Abs(New_Value <> 0) Then             'se teve alteração
            vgCtPri.Value = Abs(New_Value <> 0)                  'troca seu valor
        End If
    ElseIf vgDataType = [Option ] Then                           'campo optativo
        vgCtPri.Value = True                                     'marca esta opção
        If vgLocked = False And vgCtPri.WhatsThisHelpID <> -1 Then 'se esta travado acerta o
            vgCtFdo.Caption = Str$(BookMark)                       'label de controle do optativo
        End If
    Else                                          'campos caracteres, numéricos e memos

        'campo numérico com lista interna
        If Len(vgListFields) > 0 And (vgDataType = Numeric Or (VarType(New_Value) >= vbInteger And VarType(New_Value) <= vbCurrency)) Then
            If New_Value < vgfrmGMCale.lst.ListCount Then 'se valor passodo é menor que número de itens
                Set vgfrmGMCale.CtTransfer = Me           'classe atual
                vgfrmGMCale.lst.ListIndex = New_Value     'marca item na lst e coloca o texto no ctl principal
                vgCtPri.Text() = Substitui(vgfrmGMCale.lst.Text, vbTab, vgSubsTab$, SO_UM)
                Set vgfrmGMCale.CtTransfer = Nothing      'libera memória
            End If
        Else
            On Error Resume Next                          'previne erro
            x$ = CStr(New_Value)                          'transforma qualquer tipo em caractere
            If Err Then                                   'não conseguiu converter
                Err.Clear                                 'limpa o erro
                x$ = ""                                   'e coloca texto em branco
            End If
            vgCtPri.Text() = FormatText(x$)               'coloca o texto conforme sua máscara
        End If
    End If

    If vgSetValorOriginal Then
        vgValorOriginal = New_Value
        vgSetValorOriginal = False
    End If

DeuErro:
    If Err Then Err.Clear
    vgPriVez = vgPV
End Property

'formata uma string confome sua máscara
Public Function FormatText(St As String) As String
    Dim x As String, vgTexto As String, i As Integer, z As String, p As Integer, b10 As Double, vgNegativo As Boolean
    If IsEmpty(St) Then Exit Function             'não tem nada a fazer
    If vgDataType = Memo Then                     'campo memo
        FormatText = St$                          'nao tem formatacao
        Exit Function
    End If
    vgTexto$ = TiraTracos(St$, 1)                 'retira caracteres insertáveis do texto
    If Len(vgMascara) > 0 And vgDataType = Character And _
       Len(vgStuffChar) > 0 Then                  'ajusta texto com caracter (StuffChar) à esquerda
        vgTexto$ = Right$(String(Len(vgMascara$), vgStuffChar) + Trim$(vgTexto$), Len(TiraTracos(vgMascara$, 1))) 'com brancos
    End If
    vgTexto$ = PoeTracos(vgTexto$, 0)             'põe caracteres insertáveis no texto
    If (vgDataType = Numeric Or vgDataType = Spinner) And _
       Len(vgListFields) = 0 Then                 'lst interna com campo numérico

NovaPosDec:
        i = InStr(vgTexto$, vgSpDec$)             'testa se o separador (,) está no texto
        If i = 0 Then
            i = Len(vgTexto$) + 1                 'senão, pega o tamanho todo
        Else
            p = InStr(vgTexto$, "E-")             'verifica se número esta na base 10
            If p > 0 Then
                b10 = Val(Mid$(vgTexto$, p + 1))  'separa base 10
                If (p - 1) - i > vgQcDec Then     'trunca número de decimais
                    vgTexto$ = Left$(vgTexto$, i + vgQcDec + b10)
                End If
                b10 = ValBrasil(vgTexto$) * (10 ^ b10) 'calcula novo número
                vgTexto$ = CStr(b10)                   'coloca no texto
                GoTo NovaPosDec                        'e começa tudo de novo
            End If
        End If
        x$ = Left$(vgTexto$, i - 1)                    'separa a parte inteira do número
        If vgQcInt > 0 Then                            'trunca na qde de inteiros da máscara
            x$ = Left$(Retira$(Left$(vgTexto$, i - 1), vgSpInt$, UM_A_UM), vgQcInt)
        End If
        If vgTemSpInt Then                             'se tem separador de inteiros
            x$ = Retira$(x$, vgSpInt$, UM_A_UM)        'vamos tirar os que já tem
            vgNegativo = (Left$(Trim$(x$), 1) = "-")   'valor negativo?
            x$ = Format$(x$, "#,##0")                  'coloca sep de inteiro nos lugares certos
            If vgNegativo And Left$(x$, 1) <> "-" Then
                x$ = "-" + x$                          'o numero era negativo recoloca sinal
            End If
        End If

        'se no texto foram digitado decimais vamos colocar no texto junto com
        'o separador de decimais
        If vgQcDec > 0 Then
            z$ = Left$(Retira$(Mid$(vgTexto$, i + 1), vgSpDec$, UM_A_UM), vgQcDec)
            x$ = x$ + vgSpDec$ + Left$(z$ + String$(vgQcDec, "0"), vgQcDec)
        End If
        vgTexto$ = x$                                  'texto numérico formatado
    ElseIf vgDataType = [Date ] Then                   'cp data formata de acordo com a sua máscara
        vgTexto$ = Format$(vgTexto$, TrataMaskData$(vgMask))
    End If                                             'retorna o texto formatado
    FormatText = vgTexto$
End Function

'acha o datacontrol associado a lst externa e enche o item desta lst ou
'enche a lst interna com os itens da ListFields
Private Sub EncheRowSource()
    Dim i As Integer, z As String, k As String, QI As Integer, x As String
    On Error Resume Next                              'previne erros...
    If ComLstExt Then
        If Not vgfrmGMCale.EncheLstRs(Me) Then
            Set vgRsLista = Nothing
        Else
            Set vgfrmGMCale.CtTransfer = Me
            vgfrmGMCale.IniciaCale MOSTRA_LST_EXTERNA
        End If
        vgCtPri.Text() = vgfrmGMCale.SelecionaItem(vgValueAtual)
    ElseIf Len(vgListFields) > 0 And Len(vgBoundColumn) = 0 Then 'lst interna
        Load vgfrmGMCale                                         'carrega o form que irá conter a lst
        Set vgfrmGMCale.CtTransfer = Me                          'classe atual
        vgfrmGMCale.IniciaCale MOSTRA_LST_INTERNA                'prepara lst para ser mostrada
        z$ = vgListFields                                        'itens da lst
        vgfrmGMCale.lst.Clear                                    'limpa lista
        x$ = Substitui(vgCtPri.Text, vgSubsTab$, vbTab, SO_UM)   'item atual
        QI = -1
        Do While Len(z$) > 0                                     'vamos colocar todos os itens na lst
            i = InStr(z$, "|")                                   'próximo da lista
            If i = 0 Then i = Len(z$) + 1                        'é o último
            k$ = Trim$(Left$(z$, i - 1))                         'separa um único item
            vgfrmGMCale.lst.AddItem k$                           'adiciona na lst
            z$ = Mid$(z$, i + 1)
            If x$ = k$ Then                                      'é o item atual
                QI = vgfrmGMCale.lst.ListCount - 1               'guarda sua posição
            End If
        Loop
        If QI > -1 Then                                          'se achou o item atual
            vgfrmGMCale.lst.ListIndex = QI                       'escolhe ele na lst
        End If
        Set vgfrmGMCale.CtTransfer = Nothing                     'libera variável da memória
    End If
    DoEvents                                      'deixa o windows trabalhar
End Sub

'campo de retorno da lst externa - pega
Public Property Get BoundColumn() As String
    BoundColumn = vgBoundColumn
End Property

'campo de retorno da lst externa - coloca
Public Property Let BoundColumn(ByVal New_BoundColumn As String)
    vgBoundColumn = New_BoundColumn
End Property

'lst de itens da lst interna - pega
Public Property Get ListFields() As String
    ListFields = vgListFields
End Property

'lst de itens da lst interna - coloca
Public Property Let ListFields(ByVal New_ListFields As String)
    vgListFields = New_ListFields
    If Len(vgListFields) > 0 And Len(vgMask) > 0 Then  'cp tipo lst
        TrataMascara                                   'tem máscaras especiais
    End If
    EncheRowSource                                     'enche lst conforme novos itens
End Property

'máscara para o campo - pega
Public Property Get Mask() As String
    Mask = vgMask
End Property

'máscara para o campo - coloca
Public Property Let Mask(ByVal New_Mask As String)
    vgMask = New_Mask
    TrataMascara                                  'trata/inicializa máscara
End Property

'trata máscara/inicializa variáveis de auxilia a formatação do campo
Private Sub TrataMascara()
    Dim i As Integer, x As String                 'inicializa variáveis
    If Len(vgListFields) > 0 Then                 'se tem lst interna/externa
        vgMask = ""                               'não tem máscara
        If vgDataType <> Memo Then                'e se não for memo
            vgCtPri.MaxLength = 0                 'não tem tamanho definido
        End If
    End If
    vgMascara$ = vgMask                           'máscara tratada
    vgMaiusculo = False                           'flag - tudo maiúsculo
    vgMinusculo = False                           'flag - tudo minúsculo
    vgSoAlfa = False                              'flag - só alfabético
    vgQcDec = 0                                   'número de decimais
    vgQcInt = 0                                   'e inteiros
    vgTemSpInt = False                            'flag - tem separador de inteiro
    If Left$(vgMascara$, 1) = "I" Then            'é campo Invisível?
        vgMascara$ = Mid$(vgMascara$, 2)          'tira um caracteres da frente
    End If
    While Left$(LTrim$(vgMascara$), 1) = "@"      'enquanto começar com @ (pode combinar "@!@A...")
        vgMascara$ = LTrim$(vgMascara$)           'tira brancos da frente (pode " @! @A...")
        If Left$(vgMascara$, 2) = "@!" Then       'deseja tudo maiúsculas?
            vgMascara$ = LTrim$(Mid$(vgMascara$, 3)) 'tira dois caracteres da frente
            vgMaiusculo = True                       'liga flag do controle
        End If
        If Left$(vgMascara$, 2) = "@a" Then          'deseja tudo miúsculas?
            vgMascara$ = LTrim$(Mid$(vgMascara$, 3)) 'tira dois caracteres da frente
            vgMinusculo = True                       'liga flag do controle
        End If
        If Left$(vgMascara$, 2) = "@A" Then          'se for @A - tudo alfabético
            vgMascara$ = LTrim$(Mid$(vgMascara$, 3)) 'tira os 2 primeiros caracteres
            vgSoAlfa = True                          'liga a flag tudo alfabético
        End If
        If Left$(vgMascara$, 1) = "@" And _
           Left$(vgMascara$, 2) <> "@!" And _
           Left$(vgMascara$, 2) <> "@a" And _
           Left$(vgMascara$, 2) <> "@A" Then         'tira caracteres estranhos!!!
            vgMascara$ = LTrim$(Mid$(vgMascara$, 3))
        End If
    Wend

    'campo numérico trata máscara
    If vgDataType = Numeric Or vgDataType = Spinner Then
        x$ = Retira$(vgMascara$, ".", UM_A_UM)       'retira separação de inteiros
        i = InStr(x$, ",")                           'posição do ponto decimal
        If i > 0 Then
            vgQcInt = i - 1                          'qde inteiros a considerar
            vgQcDec = Len(x$) - i                    'qde de casas decimais a considerar
        Else                                         'não tem decimais
            vgQcInt = Len(x$)                        'só inteiros
        End If
        If InStr(vgMascara$, ".") > 0 Then           'se na máscara tem separadores de inteiros
            vgTemSpInt = True                        'seta flag
        End If
        vgMascara$ = String$(Len(vgMascara$), "9")   'coloca tudo 9 para números
    ElseIf vgDataType = [Date ] Then                 'campo data
        vgMascara$ = TrataMaskData$(vgMascara$)      'troca 99/99/99 por dd/mm/yy
        i = InStr(LCase$(vgMascara$), "mmm")         'se nome do mês é por extenso
        If i > 0 Then
            Mid$(vgMascara$, i, 1) = "!"             'troca os "m" por
            i = i + 1                                'caracteres que aceitem a digitação de alfabéticos

OutroMes:
            If LCase$(Mid$(vgMascara$, i, 1)) = "m" Then 'mais um "m"
                Mid$(vgMascara$, i, 1) = "a"             'troca este também
                i = i + 1                                'e vê o próximo caracter
                GoTo OutroMes
            End If
        End If
        For i = 1 To Len(vgMascara$)
            If InStr("dmyhsDMYHS", Mid$(vgMascara$, i, 1)) > 0 Then
                Mid$(vgMascara$, i, 1) = "9"             'os demais troca para aceitar digitação de números
            End If
        Next
    End If

    'se não tem @ na máscara arranja qde de caracteres possíveis de
    'serem digitados
    If InStr(vgMask, "@") = 0 And Len(vgMask) > 0 Then
        If Not vgCtPri Is Nothing Then
            If vgDataType <> Memo Then vgCtPri.MaxLength = Len(vgMascara$)
        End If
    End If
End Sub

'tido do campo - pega
Public Property Get DataType() As TipoDeCampo
    DataType = vgDataType
End Property

'tido do campo - coloca
Public Property Let DataType(ByVal New_DataType As TipoDeCampo)
    vgDataType = New_DataType
End Property

'caractere a justificar a esquerda - pega
Public Property Get StuffChar() As String
    StuffChar = vgStuffChar
End Property

'caractere a justificar a esquerda - coloca
Public Property Let StuffChar(ByVal New_StuffChar As String)
    vgStuffChar = New_StuffChar
End Property

'tira os caracteres insertáveis, na posição correspondente da máscara
Private Function TiraTracos(vgSt As String, vgPo As Integer) As String
    Dim k As String, i As Integer, x As String, z As String, w As String          'dimensiona locais
    z$ = C_INSERTAVEIS                            'caracteres que podem ser inseridos nas máscaras
    If vgDataType = Numeric Or vgDataType = Spinner Then
        z$ = Mid$(z$, 5)                          'tira sinais e separadores para campos numéricos
    End If
    If InStr(vgMascara$, "#") Then                'se na máscara pode ter brancos
        z$ = Left$(z$, Len(z$) - 1)               'tira o espaco dos caracteres insertáveis
    End If
    k$ = Left$(vgSt$, vgPo - 1)                   'salva o que tem antes do cursor
    For i = vgPo To Len(vgSt$)                    'com o restante,
        x$ = Mid$(vgSt$, i, 1)                    'caractere do campo
        w$ = Mid$(vgMascara$, i, 1)               'caractere da máscara
        If (InStr(z$, x$) = 0 Or w$ = "" Or (InStr("Aa!X#", w$) > 0 And x$ = " ")) And _
           Asc(Mid$(vgSt$, i)) <> 255 Then        'se não for caractere insertável...
            k$ = k$ + x$                          'concatena com o que tinha antes
        End If
    Next
    TiraTracos$ = k$                              'depois de tudo retorna a nova string
End Function

'coloca caracteres insertáveis, na posição correspondente da máscara
Private Function PoeTracos(vgSt As String, vgPo As Integer) As String
    Dim k As String, x As String, i As Integer, _
        j As Integer                                   'dimensiona locais
    k$ = Left$(vgSt$, vgPo)                            'salva o que tem antes do cursor
    j = vgPo + 1                                       'inicializa j no próximo char
    For i = vgPo + 1 To Len(vgSt$)                     'com o restante,
        x$ = Mid$(vgMascara$, j, 1)                    'pega o caracter que tem na máscara, nesta posição
        If InStr(C_INSERTAVEIS, x$) > 0 Or x$ = "" Then 'se for um char insertável,
            k$ = k$ + Mid$(vgMascara$, j, 1)            'tira da máscara e coloca no campo
            j = j + 1                                   'incrementa contador
        End If
        k$ = k$ + Mid$(vgSt$, i, 1)                     'concatena com o restante
        j = j + 1                                       'incrementa contador
    Next                                                'pega próximo
    PoeTracos$ = k$                                     'retorna string com traços
End Function

'faz o tratamento da tecla pressionada durante a edição do controle...
'como o controle de edição é todo feito aqui, vamos retornar sempre 0 para
'que o VB fique na dele e não faça nada...
Private Function TrataTecla(KeyAscii As Integer, ByVal KeyOri As Integer) As Integer
    Dim p As Integer, vgTexto As String, vgCM As String, z As String, PV As Integer, _
        x As String, i As Integer, vgTamaCp As Integer, j As Integer, vgClipB As String, _
        vgNegativo As Integer, ii As Long         'dimensiona locais

    If vgDataType <> [Boolean] And vgDataType <> [Option ] Then 'todos os tipos, menos estes...
        vgTamaCp = IIf(vgCtPri.MaxLength = 0, 32000, vgCtPri.MaxLength) 'ajusta qde de máxima de chars a digitar
    End If
    If vgSoAlfa Then                              'se máscara alfabética,
        vgMascara$ = String$(vgTamaCp, 65)        'enchemos a máscara de AAAA...
    End If

    Select Case KeyAscii                          'tecla digitada
        Case vbKeyReturn, -1                      'se ENTER ou seta para baixo
            If vgDataType = Memo Then             'e o campo for memo
                TrataTecla = KeyOri               'vamos deixar com o VB
            Else                                  'senão,
                TrataTecla = 0
                SendK vbKeyTab                    'funcionará com TAB
            End If
            Exit Function                         'cai fora daqui...

        Case -2                                   'seta para cima...
            If vgDataType = Memo Then             'nos campo memos
                TrataTecla = KeyOri               'vamos deixar com o VB
            Else                                  'nos demais
                TrataTecla = 0                    'anula tecla digitada e
                keybd_event VK_SHIFT, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0 'Simula KeyPress
                SendK vbKeyTab                                            'trata como SHIFT+TAB
                keybd_event VK_SHIFT, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0 'Simula KeyRelease
            End If
            Exit Function                         'e sai...
    End Select

    If vgLocked Or vgCtPri.WhatsThisHelpID = -1 Then 'controle não editável
        KeyAscii = 0                                 'tecla não faz nada
        Exit Function                                'cai fora...
    End If

    If vgDataType <> [Boolean] Then
        Select Case KeyAscii                         'tecla digitada
            Case 3, 24                               'CTRL+C, CTRL+X
                Clipboard.Clear                      'limpa o clipboard
                Clipboard.SetText vgCtPri.SelText    'joga texto selecionado nele
                If KeyAscii = 24 Then                'se CTRL+X
                    KeyAscii = TECLA_DEL             'mata texto marcado
                Else
                    TrataTecla = 0                   'anula tecla
                    Exit Function                    'e cair fora daqui...
                End If

            Case 22                                  'CTRL+V
                vgClipB$ = Clipboard.GetText()       'pega o que esta no clipboard
                If Len(vgMascara) = 0 Then           'se não tem máscara
                    p = vgCtPri.SelStart + 1         'poe o clipboard todo de uma vez
                    vgCtPri.Text = Left$(vgCtPri.Text, p - 1) + IIf(vgMaiusculo, UCase$(vgClipB$), IIf(vgMinusculo, LCase$(vgClipB$), vgClipB$)) + _
                                   Mid$(vgCtPri.Text, p + vgCtPri.SelLength)
                    vgCtPri.SelStart = p + Len(vgClipB$) - 1 'reposiciona posiçao do cursor
                Else                                         'tem máscara
                    If vgDataType = Numeric Or vgDataType = Spinner Then 'se campo numérico
                        vgClipB$ = Retira$(vgClipB$, vgSpInt$, SO_UM)    'tira separador de inteiros
                    End If
                    For ii = 1 To Len(vgClipB$)                          'joga letra por letra
                        KeyAscii = Asc(Mid$(vgClipB$, ii, 1))            'para tratar a máscara
                        GoSub PoeLetra
                    Next
                End If
                TrataTecla = 0                                           'retorna 0 - VB não faz nada... Tudo foi tratado aqui...
                Exit Function                                            'cai fora daqui...
        End Select
    End If

    'se tipo de campo for memo, lógico ou optativo não tem máscara
    If vgDataType = Memo Or vgDataType = [Boolean] Or vgDataType = [Option ] Then
        TrataTecla = KeyOri                                              'deixa o VB trabalhar normalmente
    Else
        GoSub PoeLetra                            'trata tecla de acordo com sua máscara
        TrataTecla = 0                            'retorna 0 - VB não faz nada... Tudo foi tratado aqui...
    End If
    Exit Function

PoeLetra:
    If vgCtPri.Locked Or vgCtPri.WhatsThisHelpID = -1 Then 'travado cai fora...
        Exit Function
    End If

    p = vgCtPri.SelStart + 1                               'pega posição após o cursor
    vgTexto$ = vgCtPri.Text                                'salva o texto que está no campo
    If vgCtPri.SelLength > 0 Then                          'se tem texto selecionado,
        vgTexto$ = Left$(vgTexto$, p - 1) + _
        String$(vgCtPri.SelLength, 255) + _
        Mid$(vgTexto$, p + vgCtPri.SelLength)              'joga asc 255 na parte selecionada
    End If
    vgCM$ = Mid$(vgMascara$, p, 1)                         'pega caractere da máscara, nesta posição
    If vgCM$ = "" Then vgCM$ = "X"                         'se x ou nada, qq caracter serve
    If HaNaString(1, vgMascara$, C_INSERTAVEIS, UM_A_UM) > 0 Or _
     InStr(vgTexto$, Chr$(255)) > 0 Then                   'se tem caractere insertável na máscara ou tinha seleção
        vgTexto$ = TiraTracos(vgTexto$, p)                 'retira insertáveis a partir do cursor
    End If
    z$ = vgTexto$
    If KeyAscii = vbKeyBack Then                           'se teclou retrocesso,
        If p > 1 Then                                      'se não estamos na primeira posição,
            vgTexto$ = Left$(vgTexto$, p - 2) + Mid$(vgTexto$, p) 'retira o caractere e
            p = p - 1                                             'reajusta posição
        End If
    ElseIf KeyAscii = TECLA_DEL Then                              'se teclou DEL
        If vgCtPri.SelLength = 0 Then
            vgTexto$ = Left$(vgTexto$, p - 1) + Mid$(vgTexto$, p + 1) 'retira o caractere sob o cursor
        End If
    ElseIf (vgDataType = Numeric Or vgDataType = Spinner) And (((KeyAscii = 44 Or KeyAscii = 46) And (InStr(vgTexto$, vgSpDec$) > 0 Or vgQcDec = 0)) Or _
      ((KeyAscii = 45 And (InStr(vgTexto$, "-") > 0 Or p > 1)))) Then 'se dado numérico, teclou vírgula e já tinha vírgula,
            Beep                                  'beep!
    Else                                          'senão,
        If KeyAscii < 32 Or KeyAscii = 127 Then   'se for char estranho
            TrataTecla = 0                        'não considera
            Exit Function                         'e cai fora daqui...
        End If
        If p <= vgTamaCp Then                     'se posição cursor dentro do campo,
            If InStr(C_INSERTAVEIS, vgCM$) > 0 And _
               vgDataType <> Numeric And vgDataType <> Spinner Then 'e se está sobre um caractere insertável
                vgTexto$ = Left$(vgTexto$, p - 1) + vgCM$ + _
                           Mid$(vgTexto$, p)                        'tira este caractere
                p = p + 1                                           'incrementa a posição
                vgCM$ = Mid$(vgMascara$, p, 1)                      'pega o novo caractere da máscara, na posição
                If vgCM$ = "" Then vgCM$ = "X"                      'se vazio, considera x - qq caractere serve
            End If
            Select Case vgCM$                                       'compara o caractere que está na máscara, na posição

                Case "9", "#"                                       'números 0-9 e sinais (se # pode espaço também)
                    x$ = C_NUMERICOS                                'inicializa x com dígitos numéricos
                    If vgCM$ = "#" Then                             'se na máscara tem # e tipo de dado não numérico
                        If vgDataType <> Numeric And vgDataType <> Spinner Then
                            x$ = x$ + " "                           'espaço também vai poder ser digitado...
                        End If
                    ElseIf vgDataType = Numeric Or vgDataType = Spinner Then 'senão, se tipo de dado é numérico,
                        x$ = x$ + "-,."                                      'vamos poder digitar vírgula e sinais
                        If KeyAscii = 44 Or KeyAscii = 46 Then
                            KeyAscii = Asc(vgSpDec$)
                        End If
                    End If
                    If InStr(x$, Chr$(KeyAscii)) Then                        'se o que teclamos for válido, então
                        vgTexto$ = Left$(vgTexto$, p - 1) + Chr$(KeyAscii) + _
                                 Mid$(vgTexto$, p)                           'concatena na string a ser retornada
                        p = p + 1                                            'incrementa a posição
                    End If
                Case "A", "N"                                                'só letras (A-z acentuadas ou não) - se "N", pode números também
                    If (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 223 And KeyAscii < 253) Or _
                       (KeyAscii > 192 And KeyAscii < 221) Or (KeyAscii > 96 And KeyAscii < 123) Or _
                       (KeyAscii = 32) Or (vgCM$ = "N" And _
                       InStr(C_NUMERICOS, Chr$(KeyAscii)) > 0) Then          'se for caracter aceitável
                        vgTexto$ = Left$(vgTexto$, p - 1) + _
                                 IIf(vgMaiusculo, UCase$(Chr$(KeyAscii)), Chr$(KeyAscii)) + _
                                 Mid$(vgTexto$, p)                           'concatena controlando maiúsculas...
                        p = p + 1                                            'incrementa ponteiro do cursor
                    End If
                Case "!", "X", "a"                                           'qualquer caracter serve...
                    vgTexto$ = Left$(vgTexto$, p - 1) + _
                             IIf(vgMaiusculo Or vgCM$ = "!", UCase$(Chr$(KeyAscii)), IIf(vgMinusculo Or vgCM$ = "a", LCase$(Chr$(KeyAscii)), Chr$(KeyAscii))) + _
                             Mid$(vgTexto$, p)                               'concatena jogando para maiúsculas, se requerido
                    p = p + 1                                                'incrementa ponteiro do cursor
            End Select
        End If
    End If
    If HaNaString(1, vgMascara$, C_INSERTAVEIS, UM_A_UM) > 0 Then            'se tem caracter insertável na máscara,
        vgTexto$ = Left$(PoeTracos(vgTexto$, p - 1), vgTamaCp)               'põe os caracteres nas posições
    End If
    If vgDataType = Numeric Or vgDataType = Spinner Then                     'campos numéricos
        i = InStr(vgTexto$, vgSpDec$) - 1                                    'testa se o separador (,) está no texto
        If i = -1 Then i = Len(vgTexto$)                                     'se não está só tem inteiros
        x$ = Left$(vgTexto$, i)                                              'pega a parte inteira
        If vgTemSpInt Then                        'se existe separador de inteiros
            j = Len(z$) - Len(Retira$(z$, vgSpInt$, UM_A_UM)) 'qde de sep inteiros que tem no texto
            x$ = Left$(Retira$(x$, vgSpInt$, UM_A_UM), vgQcInt) 'deixa só os números
            vgNegativo = (Left$(Trim$(x$), 1) = "-")            'valor negativo?
            x$ = Format$(x$, "#,##0")                           'coloca sep de inteiro nos lugares certos
            If vgNegativo And Left$(x$, 1) <> "-" Then
                x$ = "-" + x$                                   'o numero era negativo recoloca sinal
            End If
            x$ = Substitui$(x$, ".", vgSpInt$, UM_A_UM)         'coloca o separador de acordo com o idioma
            p = p + (Len(x$) - Len(Retira$(x$, vgSpInt$, UM_A_UM)) - j) 'posição atual do cursor
            If KeyAscii = 45 Then p = p + 1                             'se digitou (.) anda mais um
        Else
            x$ = Left$(x$, vgQcInt)                                     'deixa só os números da parte inteira...
        End If

        'coloca decimais e o peprador de decimais se necessário
        vgTexto$ = x$ + Left$(Retira$(Mid$(vgTexto$, i + 1), vgSpInt$, UM_A_UM), vgQcDec + 1)
    End If
    vgCtPri.Text = vgTexto$                                             'joga o resultado no controle principal
    If p <= 0 Then p = 1
    vgCtPri.SelStart = p - 1                                            'restaura a posição do cursor
    Return

End Function

'título do controle - pega
Public Property Get Caption() As String
    If Not (vgCtTit Is Nothing) Then              'se o ctl do título já foi definido
        Caption = vgCtTit.Caption                 'pega sua propriedade
    End If
End Property

'título do controle - coloca
Public Property Let Caption(ByVal New_Caption As String)
    If Not (vgCtTit Is Nothing) Then              'se o ctl existe
        vgCtTit.Caption() = New_Caption           'coloca o novo título
    End If
End Property

'tipo do shape quando sem o foco - pega
Public Property Get FillStyleOnFocus() As Integer
    FillStyleOnFocus = vgFillStyleOnFocus
End Property

'tipo do shape quando sem o foco - coloca
Public Property Let FillStyleOnFocus(ByVal New_FillStyleOnFocus As Integer)
    vgFillStyleOnFocus = New_FillStyleOnFocus
    If Not (vgCtShp Is Nothing) Then              'se o controle já está definido
        vgCtShp.FillStyle = vgFillStyleOnFocus    'troca propriedade do shape
    End If
End Property

'tipo do shape quando com o foco - pega
Public Property Get FillStyleNoFocus() As Integer
    FillStyleNoFocus = vgFillStyleNoFocus
End Property

'tipo do shape quando com o foco - coloca
Public Property Let FillStyleNoFocus(ByVal New_FillStyleNoFocus As Integer)
    vgFillStyleNoFocus = New_FillStyleNoFocus
End Property

'nome do data control da lst externa - pega
Public Property Get RowSource() As String
    RowSource = vgRowSource
End Property

'nome do data control da lst externa - coloca
Public Property Let RowSource(ByVal New_RowSource As String)
    vgRowSource = New_RowSource
End Property

'verifica se o campo tem uma lista externa
Public Function ComLstExt() As Integer
    If Not (vgRsLista Is Nothing) And Len(vgBoundColumn) > 0 And Len(vgListFields) > 0 Then
        ComLstExt = True
    Else
        ComLstExt = False
    End If
End Function

'habilita controle - pega
Public Property Get Enabled() As Boolean
    On Error Resume Next
    Enabled = vgCtPri.Enabled
End Property

'habilita controle - coloca
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    On Error Resume Next
    If Not (vgCtPri Is Nothing) Then vgCtPri.Enabled() = New_Enabled

    'verifica os controles associados e muda seu estado também
    If Not (vgCtBot Is Nothing) Then vgCtBot.Enabled() = New_Enabled
    If Not (vgCtScr Is Nothing) Then vgCtScr.Enabled() = New_Enabled
    If Not (vgCtFdo Is Nothing) Then vgCtFdo.Enabled() = New_Enabled
    If Not (vgCtTit Is Nothing) Then vgCtTit.Enabled() = New_Enabled
End Property

'visibilidade do controle - pega
Public Property Get Visible() As Boolean
    Visible = vgVisible
End Property

'visibilidade do controle - coloca
Public Property Let Visible(ByVal New_Visible As Boolean)

    On Error Resume Next
    vgVisible = New_Visible
    If Not (vgCtPri Is Nothing) Then vgCtPri.Visible() = vgVisible

    'troca visibilidade dos controles associados
    If Not (vgCtBot Is Nothing) Then vgCtBot.Visible() = vgVisible
    If Not (vgCtScr Is Nothing) Then vgCtScr.Visible() = vgVisible
    If Not (vgCtFdo Is Nothing) And vgDataType <> [Option ] Then vgCtFdo.Visible() = vgVisible
    If Not (vgCtTit Is Nothing) Then vgCtTit.Visible() = vgVisible
    If Not (vgCtShp Is Nothing) Then vgCtShp.Visible() = vgVisible
End Property

'nome do controle - pega
Public Property Get Name() As String
    Name = vgCtPri.Name
End Property

'indice do controle - pega
Public Property Get Index() As Integer
    Index = vgCtPri.Index
End Property

'inicializa classe com os defaults de algumas propriedades
Private Sub Class_Initialize()
    vgColorNoFocus = &HFF&
    vgColorOnFocus = &HFF00&
    vgFillStyleOnFocus = 0
    vgFillStyleNoFocus = 0
    vgFormatLostFocus = True
    vgBookMark = Null
    vgMask = ""
    vgListFields = ""
    vgBoundColumn = ""
    vgDataType = Character
    vgStuffChar = ""
    vgSelectOnFocus = True
    vgLocked = False
    vgfrmCalcNaMem = False
    vgSetValorOriginal = False
    vgValueAtual = Empty
End Sub

'termina a classe
Private Sub Class_Terminate()
    On Error Resume Next                          'previne erros
    If Len(vgListFields) > 0 Then                 'se tem lst
        Unload vgfrmGMCale                        'descarrega o form da lst
    End If
    If vgfrmCalcNaMem Then                        'se a calculadora está na memória
        Unload frmGMCalc                          'vamos descarregá-la
    End If
End Sub


