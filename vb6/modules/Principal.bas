Attribute VB_Name = "Principal"
'*******************************************************************
' Projeto    : Hi-Tech - Sistema Integrado de Gestão Comercial
' Versão     : 2.6
' Módulo     : Principal
' Autor      : Fabiano
' Modificação: 03/05/2010
' Objetivo   : Este módulo possui as rotinas gerais do sistema
'*******************************************************************
Option Explicit

Public cnxAdmn As String      'Conexão ao bd xn_adm
Public cnxData As String      'Conexão ao bd xn_dbf
Public cnxCEPs As String      'Conexão ao bd cep
Public cnxMail As String      'Conexão ao bd e-mails
Public cnxRela As String      'Conexão ao bd relatórios
'Public gdb As Database        'Conexão aos bds

Public dbAdmn As Database     'Conexão ao bd xn26_adm
Public dbData As Database     'Conexão ao bd xn26_dbf
Public dbRela As Database     'Conexão ao bd xn26_rel
Public dbCeps As Database     'Conexão ao bd cep
Public dbMail As Database     'Conexão ao bd e-mails

Public appPathApp As String   'Armazena o diretório do programa
Public appPathIni As String   'Armazena o local do arquio ini
Public appPathRpt As String   'Armazena o diretório de relatórios
Public appPathLbl As String   'Armazena o diretório do rótulos
Public appPathNFe As String   'Armazena o diretório das NF-e
Public appEXEName As String   'Armazena o nome do executável do programa
Public appIDEmpresa As String 'Armazena o ID da empresa para validação da licença
Public appLicenca As String   'Armazena a licença de uso do aplicativo
Public appURLUpdt As String   'Armazena o local de atualização

Public ImpressoraPadrao As String  'Armazena o nome da impressora padrão

'Variáveis da configuração das impressoras dos rótulos
Public imprRotPA As String    'Produto acabado
Public imprRotCQ As String    'Controle de qualidade
Public imprRotMP As String    'Matéria-prima
Public imprRotCM As String    'Controle qualidade matéria-prima
Public imprRotLS As String    'Produto acabado (lateral sacaria)
Public imprRotAM As String    'Rótulo amostra (volumes expedição)

'Variáveis das opções gerais
Public expdBox As Integer     'N.° do box da expedição

'Public BigIcons As ImageList  'Variável do controle imlIcons
'Public SmlIcons As ImageList  'Variável do controle imlSmlIcons
Public imgIcons1 As Collection   'Coleção de ícones
Public imgIcons2 As Collection

Public cLog As Logon          'Classe do sistema de login
Public cRpt As CrystalReport

'Public cRpt As CrystalReport  'Variável do control crpRel

'Public fMain As frmMain          'Formulário Principal

Public FormParent As String      'Formulário de origem
Public CloseSystem As Boolean    'O sistema está sendo fechado

Public GroupPrograms As Collection  'Grupos de programas
Public GroupIcons As Collection     'Icones dos grupos de programas

'Constantes utilizadas no projeto
Public Const xnPrjName = "Hi-Tech"              'Nome do projeto
Public Const xnODBCName = "hi-tech v2.6"        'Nome fonte de dados ODBC
Public Const xnODBCRela = "hi-tech reports v2.6"   'Nome fonte de dados dos relatórios temporários
Public Const xnArqvINI = "xn.ini"               'Nome do arquivo de configurações
Public Const xnArqvRes = "xnres.dll"            'Nome do arquivo de recursos

'Constantes para formatação
Public Const xnMONEY = "###,###,###,##0.00"     'Números
Public Const xnMONEY4 = "###,###,###,##0.0000"  'Números
Public Const xnPESO = "###,###,###,##0.000"     'Peso de produtos
Public Const xnLOTE = "0000-00000"              'Lote de produtos
Public Const xnLAUDO = "00000/00000-0"          'Laudo de análise

Public Const xnKEYTAB = 9                       'Caracter TAB
Public Const xnCNPJ = "00\.000\.000/0000-00"    'CNPJ
Public Const xnCPF = "000\.000\.000-00"         'CPF
Public Const xnCEP = "00000-000"                'CEP
Public Const xnPHONE = "(00)0000-0000"          'telefone/fax
Public Const xnPLACA = "@@@-@@@@"               'placa
Public Const xnDATA = "dd/mm/yyyy"              'data
Public Const xnHORA = "hh:nn:ss"                'hora
Public Const xnHRMN = "hh:nn"                   'hora em minuto
Public Const xnDTHR = "dd/mm/yyyy hh:nn:ss"     'data e hora
Public Const xnDTHM = "dd/mm/yyyy hh:nn"

Public Const xnMASK = &HFF00&
Public Const xnTAGPACK = "EMBALAR"  'Tag para embalagem nas OP's

'Constantes gerais das classes
Public Const HRS_TRAB_DIA As Long = 528

'Constantes para compatibilidade da contabilidade
Public Const FIN_LCONTABIL_AF = 3234   'Código contábil Cheques
Public Const FIN_LCONTABIL_PR = 99999  'Código contábil de pagamento parcial do AF
Public Const FIN_LCONTABIL_JR = 603    'Código contábil juros
Public Const FIN_LCONTABIL_DC = 804    'Código contábil desconto

'Constantes para e-mail
Public Const xnFOLDER_ROOT = "ROOT"         'Pastas locais
Public Const xnFOLDER_ENTR = "INBOX"        'Caixa de entrada
Public Const xnFOLDER_SAID = "OUTBOX"       'Caixa de saída
Public Const xnFOLDER_IENV = "ITEMSSEND"    'Itens enviados
Public Const xnFOLDER_IEXC = "ITEMSDEL"     'Itens excluídos

'Constantes para modo de edição dos registros nos módulos Classes
Public Const modInclude = "Inclusão"
Public Const modEdit = "Alteração"
Public Const modQuery = "Consulta"
Public Const modClosed = "Fechado"
Public Const modView = "Visualizar"    'Incluído v2.6

'Constantes para agendamento do backup
Public Const bkpHORA1 = #12:00:00 AM#     'Meia noite
Public Const bkpHORA2 = #6:00:00 AM#      '6 da manhã
Public Const bkpHORA3 = #12:15:00 PM#     'Meio dia
Public Const bkpHORA4 = #6:00:00 PM#      '6 da tarde

Public Const bkpInicio = 3     'Tempo limite antes do backup
Public Const bkpEspera = 15    'Tempo utilizado para o backup

'Mensagens do programa /////////////////////////////////
Public Const Msg0001 = "O aplicativo já está sendo executado."
Public Const Msg0002 = "Não foi possível estabelecer uma conexão com o banco de dados."
Public Const Msg0003 = "O programa não está disponível para execução no sistema." & vbCr & "Contate o administrador do sistema."
Public Const Msg0004 = "Usuário sem permissão de executar este programa."

Public Const Msg1003 = "Número do erro: "
Public Const Msg1006 = "Para obter Ajuda, pressione F1."
Public Const Msg1007 = "Tem certeza que deseja finalizar o programa ?"

Public Const Msg1202 = "Se o registro for excluído não poderá ser recuperado."
Public Const Msg1203 = "Tem certeza que deseja excluir o registro ?"
Public Const Msg1204 = "Esta ação cancelará todas as alterações feitas no registro."
Public Const Msg1205 = "Deseja salvar as alterações no registro atual ?"

Public Const Msg1602 = "Nenhum registro foi selecionado."

'Declarações API
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'*******************************************************************
' Procedimento: Main
' Argumentos  : Nenhum
' Retorno     : Nenhum
' Objetivo    : Procedimento de inicialização do sistema
'*******************************************************************
Public Sub Main()
   ChDir App.Path             'Muda o diretório padrão para onde está o sistema
   appPathApp = App.Path      'Armazena o diretório do sistema
   NormalizePath appPathApp   'Normaliza o diretório
   appPathRpt = appPathApp & "relatorios\"   'Armazena o diretório dos relatórios
   appPathLbl = appPathApp & "rotulos\"      'Armazena o diretório dos rótulos
   appPathIni = appPathApp & xnArqvINI       'Armazena o arquivo ini
   appEXEName = App.EXEName & ".exe"
   
   'HoraServidor cSvr.Alias    'Sincroniza a hora com a do servidor
   
   'Se True, não permite a execução do sistema pois está no horário
   'de execução do backup do sistema
   'If HoraBackup(True) Then
      'KillApp appEXEName
   '   Exit Sub
   'End If
   
   'Armazena as configurações do sistema
   LerConfiguracao
   
   'Carrega o formulário principal
   On Local Error Resume Next
   PrepareThemeSupport
   
   'Set fMain = New frmMain
   frmMain.Show
End Sub

'*******************************************************************
' Procedimento: HoraBackup
' Argumentos  : ExibirMsg As Booelan
'               -> Se True, então será mostrado uma mensagem de alerta
' Retorno     : Boolean
'               Se True, avisa sobre a operação de backup e não
'               executa o sistema
'               Se False, continuar a execução do sistema
' Objetivo    : Verificar se a hora atual está no intervalo agendado
'               para a realização do backup das bases de dados
'*******************************************************************
Function HoraBackup(ByVal ExibirMsg As Boolean) As Boolean
   Dim hIni As Date, hFim As Date      'Declara as variáveis
   Dim sMsg As String
   
   'Subtrai da hora atual o tempo para inicio do backup
   hIni = Format$(DateAdd("n", -bkpInicio, bkpHORA1), xnHORA)
   'Adiciona a hora atual o tempo gasto para o backup
   hFim = Format$(DateAdd("n", bkpEspera, bkpHORA1), xnHORA)
   
   'Compara se a hora atual está dentro do intervalo agendado para o
   'backup. Em caso positivo passa a mensagem para a variável sMsg e
   'pula para a linha IniBackup
   If (Time >= hIni) And (Time <= hFim) Then
      sMsg = "O programa não pode ser iniciado devido ao backup agendado para às 00:00 h."
      GoTo IniBackup
   End If
   
   'Repete o processo para cada hora agendada
   hIni = Format$(DateAdd("n", -bkpInicio, bkpHORA2), xnHORA)
   hFim = Format$(DateAdd("n", bkpEspera, bkpHORA2), xnHORA)
   
   If (Time >= hIni) And (Time <= hFim) Then
      sMsg = "O programa não pode ser iniciado devido ao backup agendado para às 06:00 h."
      GoTo IniBackup
   End If
   
   hIni = Format$(DateAdd("n", -bkpInicio, bkpHORA3), xnHORA)
   hFim = Format$(DateAdd("n", bkpEspera, bkpHORA3), xnHORA)
   
   If (Time >= hIni) And (Time <= hFim) Then
      sMsg = "O programa não pode ser iniciado devido ao backup agendado para às 12:15 h."
      GoTo IniBackup
   End If
   
   'Removido o agendamento das 18h
   'hIni = Format$(DateAdd("n", -bkpInicio, bkpHORA4), xnHORA)
   'hFim = Format$(DateAdd("n", bkpEspera, bkpHORA4), xnHORA)
   
   'If (Time >= hIni) And (Time <= hFim) Then
   '   sMsg = "O programa não pode ser iniciado devido ao backup agendado para às 18:00 h."
   '   GoTo IniBackup
   'End If
   
   HoraBackup = False   'Retona False, pois a hora não está no intervalo agendado
   Exit Function        'Sai da função
   
IniBackup:
   'Se o argumento é True mostra uma mensagem de alerta
   If ExibirMsg Then ShowMsg sMsg, vbInformation
   'Retora True, pois a hora está dentro do intervalo agendado para o backup
   HoraBackup = True
End Function

'*******************************************************************
' Procedimento: MsgInfo
' Argumentos  : Msg As String
'               -> Passa um texto para exibição
' Retorno     : Nenhum
' Objetivo    : Esta rotina atualiza o status do andamento no form
'               Splash
'*******************************************************************
Public Sub MsgInfo(ByVal Msg As String, Optional ByVal Pause As Long = 0)
   frmMain.lblMsg = Msg
   frmMain.lblMsg.Refresh
   If Pause > 0 Then Sleep Pause   'Realiza uma pausa
End Sub

'*******************************************************************
' Procedimento: LerConfiguração
' Argumentos  : Nenhum
' Retorno     : Nenhum
' Objetivo    : Recupera a configuração do sistema armazena no
'               registro do Windows
'*******************************************************************
Public Sub LerConfiguracao()
   Dim vValue As Variant      'Declara as variáveis
   Dim lDC As Long
   Dim cIni As Ini
   
   'Inicializa o objeto de controle de arquivos INI
   Set cIni = New Ini
   
   'Seta o nome do arquivo
   cIni.Arquivo = appPathIni
   
   'Recupera a licenca de uso do programa
   vValue = cIni.LerTexto("LICENCA", "CompanyID", "")
   appIDEmpresa = vValue
   
   vValue = cIni.LerTexto("LICENCA", "Key", "")
   appLicenca = vValue
   
   'Recupera o local das pastas NF-e
   vValue = cIni.LerTexto("NFE", "PastaNFe", "\\HI-TECH02\DOCS\18-FISCAL\")
   appPathNFe = vValue
   
   'Recupera a configuração das impressoras do rótulo
   vValue = cIni.LerTexto("ROTULOS", "ProdutoAcabado")
   imprRotPA = IIf(IsEmpty(vValue), "", vValue)
   
   vValue = cIni.LerTexto("ROTULOS", "ControleQualidade")
   imprRotCQ = IIf(IsEmpty(vValue), "", vValue)
   
   vValue = cIni.LerTexto("ROTULOS", "MateriaPrima")
   imprRotMP = IIf(IsEmpty(vValue), "", vValue)
   
   vValue = cIni.LerTexto("ROTULOS", "CQMateriaPrima")
   imprRotCM = IIf(IsEmpty(vValue), "", vValue)
   
   vValue = cIni.LerTexto("ROTULOS", "PALateralSacaria")
   imprRotLS = IIf(IsEmpty(vValue), "", vValue)
   
   vValue = cIni.LerTexto("ROTULOS", "AmostraExpedicao")
   imprRotAM = IIf(IsEmpty(vValue), "", vValue)
   
   'Recupera a configuração de atualização
   vValue = cIni.LerTexto("GERAL", "URLAtualizacao", "\\HI-TECH02\PUBLICA\SOFTWARE\")
   appURLUpdt = vValue
   
   'Recupera a configuração do box da expedição
   vValue = cIni.LerTexto("GERAL", "ExpedicaoBox", 0)
   expdBox = Val(vValue)
   
   'Recupera a configuração da impresora
   vValue = cIni.LerTexto("GERAL", "ImpressoraPadrão")
   
   'Verifica se a configuração é válida
   If IsEmpty(vValue) Then
      Dim bRet As Boolean  'Declara as variáveis
      Dim cPrt As Object
      
      'Exibe mensagem informando que não há impressora configurada
      ShowMsg "Não há uma impressora definida como padrão para uso do sistema." & vbCr & _
         "Por favor selecione uma impressora padrão para continuar.", vbInformation
      
      'Mostra a janela de Configuração da Impressora
      'bRet = PrintDialog(lDC, , , , , , , , , , , , , cPrt, &H40)
      'Armazena o nome da impressora
      'vValue = Printer.DeviceName
      
      ImpressoraPadrao = vValue  'Transfere a configurção para a variável
      GravarConfiguracao         'Grava o novo valor da configuração
   End If
   
   'Destrói o objeto
   Set cIni = Nothing
End Sub

'*******************************************************************
' Procedimento: GravarConfiguração
' Argumentos  : Nenhum
' Retorno     : Nenhum
' Objetivo    : Grava a configuração do sistema no registro do Windows
'*******************************************************************
Public Sub GravarConfiguracao()
   Dim cIni As Ini
   
   'Inicializa o objeto de controle de arquivos INI
   Set cIni = New Ini
   
   'Seta o nome do arquivo
   cIni.Arquivo = appPathIni
   
   'Grava os novos valores das configurações gerais
   cIni.EscreverTexto "GERAL", "ImpressoraPadrão", ImpressoraPadrao
   cIni.EscreverTexto "GERAL", "URLAtualizacao", appURLUpdt
   cIni.EscreverTexto "GERAL", "ExpedicaoBox", CStr(expdBox)
   
   'Grava os novos valores da pasta da nfe
   cIni.EscreverTexto "NFE", "PastaNFe", appPathNFe
   
   'Grava os valores para as impressoras do rótulo
   cIni.EscreverTexto "ROTULOS", "ProdutoAcabado", imprRotPA
   cIni.EscreverTexto "ROTULOS", "ControleQualidade", imprRotCQ
   cIni.EscreverTexto "ROTULOS", "MateriaPrima", imprRotMP
   cIni.EscreverTexto "ROTULOS", "CQMateriaPrima", imprRotCM
   cIni.EscreverTexto "ROTULOS", "PALateralSacaria", imprRotLS
   cIni.EscreverTexto "ROTULOS", "AmostraExpedicao", imprRotAM
   
   'Destrói o objeto
   Set cIni = Nothing
End Sub

'*******************************************************************
' Procedimento: ShowAboutDialog
' Argumentos  : IconPict As StrPicutre
'               -> Ícone da janela
' Retorno     : Nenhum
' Objetivo    : Exibe a janela Sobre e mostra as informações do sistema
'*******************************************************************
Public Sub ShowAboutDialog(Optional IconPict As StdPicture = Nothing)
   Dim sVer As String      'Declara as variáveis
   Dim sLicense As String, sSerial As String
   Dim pic As StdPicture
   
   'Passa os valores para as variáveis
   Set pic = IIf((IconPict Is Nothing), frmMain.Icon, IconPict)
   sVer = VersaoPrograma  'App.Major & "." & App.Minor & "." & App.Revision
   sLicense = "" 'GetLicenseToApp(xnPrjName)
   sSerial = "" 'GetSerialNumberToApp(xnPrjName)
   
   'Execute a rotina que exibe a janela
   'AboutBox xnPrjName, sVer, App.LegalCopyright, pic, sLicense, sSerial
End Sub

'*******************************************************************
' Procedimento: ShowHelpDialog
' ArgumentoS  : FormOrig As Long
'               -> Handle do formulário de onde será executado a rotina
'               HelpID As Long
'               -> Número do tópico a ser exibido
' Retorno     : Nenhum
' Objetivo    : Executar a ajuda do sistema num determinado tópico
'*******************************************************************
Public Sub ShowHelpDialog(FormOrig As Long, Optional HelpID As Long)
   Dim nRet As Long  'Declara a variável
   
   'Execute a rotina e mostra a ajuda do programa
   'nRet = HelpDialog(xnPrjName, FormOrig, "xn.hlp", &H0, HelpID)
End Sub

'*******************************************************************
' Procedimento: ShowMsg
' Argumentos  : Prompt As String
'               -> Mensagem que será exibida
'               Buttons As Integer
'               -> Botões e ícones que serão mostrados na mensagem
' Retorno     : Integer
' Objetivo    : Mostar a MsgBox personalizada com o título do programa
'               e retornar qual ação foi escolhida pelo usuário
'*******************************************************************
Public Function ShowMsg(Prompt As String, Buttons As Integer) As Integer
   ShowMsg = MsgBox(Prompt, Buttons, xnPrjName)
End Function

'*******************************************************************
' Procedimento: IniciarPrograma
' Argumentos  : ExibirStatus As Boolean
'               -> Se True, mostrar o status de cada operação realizada
' Retorno     : Boolean
' Objetivo    : Exectua a sequencia de inicialização do programa,
'               retorna True caso não ocorra erros e False no caso
'               de alguma falha
'*******************************************************************
Public Function IniciarPrograma(ExibirStatus As Boolean) As Boolean
   'Inicia o controle de erro
   On Local Error GoTo errHandle
   
   'Exibe mensagem de andamento
   If ExibirStatus Then MsgInfo "Estabelecendo conexão ao servidor de banco de dados..."
   
   'Abre a conexão com os bancos de dados, em caso de falha
   'exibe uma mensagem de alerta e finaliza o sistema
   If Not AbrirConexaoBD Then
      ShowMsg Msg0002, vbCritical
      End
      Exit Function
   End If
   
   'Exibe mensagem de andamento
   If ExibirStatus Then MsgInfo "Conectado ao servidor de banco de dados"
   
   'Realiza uma pausa de 1 segundo
   Sleep 1000
   
   'Exibe mensagem de andamento
   'If ExibirStatus Then MsgInfo "Verificando usuário e senha..."
   
   IniciarPrograma = True     'Retorna resultado da função
   Exit Function              'Sai da função
   
errHandle:
   'Retorna resultado de erro
   IniciarPrograma = False
   
   'Gera um evento de erro no sistema e exibe para o usuário
   MsgErro "Módulo Principal", "IniciarPrograma", Err.Number, Err.Description, Erl, vbCritical, 2, , False
End Function

'*******************************************************************
' Procedimento: EncerrarPrograma
' Argumentos  : Nenhum
' Retorno     : Nenhum
' Objetivo    : Executa a sequencia de encerramento do programa,
'               fecha todas as conexões aos bancos de dados e
'               remove da memória todas as variáveis
'*******************************************************************
Public Sub EncerrarPrograma()
   'Inicia o controle de erro
   On Local Error Resume Next
   Dim i As Integer
   
   CloseThemeSupport
   
   'Finaliza o chat
   CloseSystem = True
   For i = Forms.Count - 1 To 1
      Unload Forms(i)
   Next
   
   'Fecha o form principal
   Unload frmMain
   'Set fMain = Nothing
   
   'Verifica se foi inicializada a variável do login,
   'se True executa o logout do usuário no sistema
   If Not cLog Is Nothing Then cLog.Logout
   Set cLog = Nothing      'Finaliza a variável
   
   'Verifica se as conexões foram criadas e estabelecidas,
   'se True fecha todas
   dbAdmn.CloseConnection
   dbData.CloseConnection
   dbRela.CloseConnection
   dbMail.CloseConnection
   
   'Finaliza todas as variáveis
   Set dbAdmn = Nothing
   Set dbData = Nothing
   Set dbRela = Nothing
   Set dbMail = Nothing
   
   Set imgIcons1 = Nothing
   Set imgIcons2 = Nothing
   
   Set GroupPrograms = Nothing
   Set GroupIcons = Nothing
   
   'KillApp appEXEName
End Sub

'*******************************************************************
' Procedimento: AbrirConexaoBD
' Argumentos  : Nenhum
' Retorno     : Boolean
' Objetivo    : Abre a conexão com as bases de dados e retorna True
'               em caso de sucesso ou False no caso de uma falha
'*******************************************************************
Public Function AbrirConexaoBD() As Boolean
1   On Local Error GoTo errHandle   'Inicia o controle de erro
2   Dim cn1 As String, cn2 As String
3
4   'Atribui falha na execução
5   AbrirConexaoBD = False
6
7   cn1 = "Provider=sqloledb;Network Library=DBMSSOCN;Data Source=" & cLog.Server.Name & "," & cLog.Server.Port & ";Initial Catalog=mail;User ID=" & cLog.User.Name & ";Password=" & cLog.User.Password & "Connect Timeout=10;General Timeout=0;Packet Size=8192"
8   cn2 = "Provider=sqloledb;Network Library=DBMSSOCN;Data Source=" & cLog.Server.Name & "," & cLog.Server.Port & ";Initial Catalog=xn26_rel;User ID=" & cLog.User.Name & ";Password=" & cLog.User.Password & ";Connect Timeout=10;General Timeout=0;Packet Size=8192"
9
10   cnxAdmn = cLog.AdminConnection    'Conexão bd admin
11   cnxData = cLog.DataConnection     'Conexão bd dados
12   cnxCEPs = ""                      'Conexão bd ceps
13   cnxMail = cn1                     'conexão bd mail
14   cnxRela = cn2                     'conexão bd relatórios
15
16   'Instancia os objetos
17   Set dbAdmn = New Database
18   Set dbData = New Database
19   Set dbRela = New Database
20   Set dbMail = New Database
21
22   'Abre as conexões com os bancos de dados, em caso de erro sai da função
23   If Not dbAdmn.OpenConnection(cnxAdmn) Then Exit Function
24   If Not dbData.OpenConnection(cnxData) Then Exit Function
25   If Not dbRela.OpenConnection(cnxRela) Then Exit Function
26   'If Not dbMail.OpenConnection(cnxMail) Then Exit Function
27
28   AbrirConexaoBD = True    'Conexão estabelecida
29   Exit Function            'Sai da função
30
errHandle:
31   'Gera um evento de erro no sistema e grava para posterior debug
32   MsgErro "Módulo Principal", "AbrirConexaoBD", Err.Number, Err.Description, Erl, vbCritical, 2, True, False
33   'Conexão não estabelecida
34   AbrirConexaoBD = False
End Function

'*******************************************************************
' Procedimento: InicializarCystralReports
' Argumentos  : WindowTitle As String
'               -> Título da janela do relatório
' Retorno     : Nenhum
' Objetivo    : Reseta todas as propriedades/configurações do controle
'               CrystalReports
'*******************************************************************
Public Sub InicializarCrystalReports(WindowTitle As String)
   With cRpt
      .Reset
      .DiscardSavedData = True
      .WindowTitle = WindowTitle
      .WindowLeft = 0
      .WindowTop = 0
      .WindowHeight = Screen.Height / Screen.TwipsPerPixelY
      .WindowWidth = Screen.Width / Screen.TwipsPerPixelX
      .WindowState = crptMaximized
      
      .WindowControlBox = False
      .WindowShowCancelBtn = True
      .WindowShowCloseBtn = True
      .WindowShowExportBtn = True
      .WindowShowGroupTree = False
      .WindowShowNavigationCtls = True
      .WindowAllowDrillDown = False
      .WindowShowZoomCtl = True
      .WindowShowRefreshBtn = False
      .WindowShowSearchBtn = False
      .WindowShowPrintBtn = True
      .WindowShowPrintSetupBtn = True
      .WindowShowProgressCtls = False
      
      .CopiesToPrinter = 1
      .PrinterName = ImpressoraPadrao
   End With
End Sub

'*******************************************************************
' Procedimento: MsgErro
' Argumentos  : Modulo As String
'               -> Arquivo onde gerou o erro
'               Funcao As String
'               -> Rotina onde gerou o erro
'               Numero As Long
'               -> Número do erro
'               Descricao As String
'               -> Descrição do erro
'               Linha As Integer
'               -> Linha do programa onde originou o erro
'               Opcoes As Integer
'               -> Opcoes da mensagem
'               Exibir As Boolean
'               -> Exibe a mensagem do erro
'               SalvarLog As Integer
'               -> True: salva um log do erro; False: não salva o log
' Retorno     : Nenhum
' Objetivo    : Exibe uma mensagem de erro/aviso padronizada mostrando
'               informações completas e descritivas do erro.
'*******************************************************************
Public Sub MsgErro(ByVal vModulo As String, ByVal vFuncao As String, ByVal vNumero As Long, ByVal vDescricao As String, ByVal vLinha As Integer, ByVal vOpcoes As Integer, ByVal vTipo As Integer, Optional ByVal vExibir As Boolean = True, Optional ByVal vSalvarLog As Boolean = True)
   On Local Error Resume Next
   If vExibir Then
      ShowMsg "!!! ATENÇÃO !!!" & vbNewLine & vbNewLine & _
         "Usuário:" & vbTab & cLog.User.Name & vbNewLine & _
         "Estação:" & vbTab & cLog.Workstation.ID & vbNewLine & _
         "Data:" & vbTab & Format$(Now, xnDATA) & vbNewLine & _
         "Hora:" & vbTab & Format$(Now, xnHORA) & vbNewLine & _
         "Projeto:" & vbTab & xnPrjName & vbNewLine & vbNewLine & _
         "Módulo:" & vbTab & vbTab & vModulo & vbNewLine & _
         "Procedimento:" & vbTab & vFuncao & vbNewLine & _
         "Linha:" & vbTab & vbTab & vLinha & vbNewLine & _
         "Número:" & vbTab & vbTab & vNumero & vbNewLine & _
         "Descrição:" & vbTab & vDescricao & vbNewLine, vOpcoes
   End If
   
   If vSalvarLog Then
      cLog.SaveErr cLog.User.Name, Now, cLog.Workstation.ID, xnPrjName, vModulo, vFuncao, vNumero, vDescricao, vLinha, vOpcoes, vTipo, False
   End If
   
   On Error GoTo 0
   Err.Clear
End Sub

'*******************************************************************
' Procedimento: GravarLog
' Argumentos  : vAcao As String
'               -> Ação gerada pelo sistema
' Retorno     : Nenhum
' Objetivo    : Gravar log de atividades do sistema para futuras
'               consultas e diganósticos de problemas
'*******************************************************************
Public Sub GravarLog(ByVal vAcao As String, ByVal vInfoExtra As String)
   On Local Error Resume Next
   If Not cLog Is Nothing Then
      cLog.CreateLog cLog.User.Name, Now, cLog.Workstation.ID, vAcao, vInfoExtra
   End If
End Sub

Public Sub GravarModificacaoRegistro(ByVal vAcao As String, ByVal vRegistro As Collection)
   On Local Error Resume Next
   Dim fVar As FieldChanged
   Dim sInfoExtra As String
   Dim i As Integer
   
   sInfoExtra = ""
   For i = 1 To vRegistro.Count
      Set fVar = vRegistro(i)
      sInfoExtra = sInfoExtra & "CAMPO: [" & fVar.Name & "] DE: [" & fVar.OldValue & "] PARA: [" & fVar.NewValue & "]" & vbCrLf
   Next
   
   If Not cLog Is Nothing Then
      cLog.CreateLog cLog.User.Name, Now, cLog.Workstation.ID, vAcao, sInfoExtra
   End If
End Sub

'*******************************************************************
' Procedimento: VerificarAtualizacao
' Argumentos  : Nenhum
' Retorno     : Long
' Objetivo    : Analisa o script de atualização e verifica se há
'               alguma atualização para o sistema ou arquivos.
'               Retorna o número de atualizações disponíveis
'*******************************************************************
Public Function ValidateVersion(ByVal myVer As String, ByRef newVer As String) As Boolean
   On Local Error GoTo errHandle
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim v(1 To 4) As Long
   Dim c(1 To 4) As Long
   
   'Inicializa as variáveis
   Erase v, c
   
   'Consulta a versão atual do sistema
   sSQL = "SELECT * FROM version;"
   Set r = dbAdmn.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      v(1) = r("v_major")
      v(2) = r("v_minor")
      v(3) = r("v_revision")
      v(4) = r("v_build")
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   On Error Resume Next
   'Atribui a versão do executável
   c(1) = Split(myVer, ".")(0)
   c(2) = Split(myVer, ".")(1)
   c(3) = Split(myVer, ".")(2)
   c(4) = Split(myVer, ".")(3)
   
   'Retorna o resultado
   newVer = v(1) & "." & v(2) & "." & v(3) & "." & v(4)
   ValidateVersion = (c(1) >= v(1)) And (c(2) >= v(2)) And (c(3) >= v(3)) And (c(4) >= v(4))
   Exit Function
   
errHandle:
   Debug.Print Err.Number & vbTab & Err.Description
End Function

Public Sub ExecutarAtualizacao(ByVal versao As String)
   On Local Error Resume Next
   Dim fArqv As String
   
   'Cria o nome do arquivo
   fArqv = appURLUpdt & "SetupXN_v" & versao & ".exe"
   
   'Executa o arquivo
   Shell fArqv, vbNormalFocus
End Sub

'Private Function GetComputerVerInfo() As Long
'   Dim sSQL As String
'   Dim r As ADODB.Recordset
'   Dim lRet As Long
'
'   Set r = New ADODB.Recordset
'   sSQL = "SELECT VersaoInstalada FROM Terminal WHERE (IP = '" & cLog.Workstation.ID & "');"
'   r.Open sSQL, cnxAdmn, 3
'   If Not r.BOF Then lRet = r("VersaoInstalada")
'   r.Close
'   Set r = Nothing
'
'   GetComputerVerInfo = lRet
'End Function

'*******************************************************************
' Procedimento: CriarLista
' Argumentos  : Lista As Object
'               -> O objeto pode ser uma ListBox ou ComboBox
'               Pesquisa As Long
'               -> Comando de pesquisa em SQL
'               CampoExibicao
'               -> Campo utilizado para criação da lista
'               CampoIndice
'               -> Campo utilizado para criação do índice dos registros
' Retorno     : Nenhum
' Objetivo    : Criar uma consulta no banco de dados e preencher uma
'               lista, utilizando como campo de retorno o informado
'               na variável CampoExibicao, com o valor da propriedade
'               ItemData, informado pela variável CampoIndice
'*******************************************************************
Public Sub CriarLista(Lista As Object, Pesquisa As String, CampoExibicao As String, CampoIndice As String)
   Dim sSQL As String         'Declara as variáveis
   Dim r As ADODB.Recordset
   
   Set r = dbData.OpenRecordset(Pesquisa)  'Abre a tabela
   
   Screen.MousePointer = vbHourglass   'Altera o cursor mouse para a ampulheta
   Lista.Clear                         'Limpa a lista
   
   'Percorre toda a tabela até o final
   Do While Not r.EOF
      'Adiciona o registro na lista
      Lista.AddItem r(CampoExibicao)
      
      'Armazena o índice do registro na lista
      If CampoIndice <> "" Then Lista.ItemData(Lista.NewIndex) = r(CampoIndice)
      
      'Move para o próximo registro
      r.MoveNext
   Loop
   
   On Local Error Resume Next       'Em caso de erro continua na linha seguinte
   If r.State <> 0 Then r.Close     'Fecha a tabela
   Set r = Nothing                  'Remove a variável da memória
   Screen.MousePointer = vbDefault  'Altera o cursor do mouse para o padrão
End Sub
