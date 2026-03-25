Attribute VB_Name = "General"
Option Explicit
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public dbData As Database        'Referencia a classe Database para manipulação de todo o acesso a dados

Public sysConfig As Collection   'Coleção com as configurações globais do sistema
Public maqConfig As Collection   'Coleção com as configurações locais de cada máquina
Public vTipoEdicao As String    'tipo de cadastro de produtos e tb quando abrir uma venda/editar

'1.caixa e fluxo do caixa
Public varCodCaixa As Long         'pegar o codigo do caixa no fluxo
Public varFluxoCaixa As Boolean    'saber de onde foi acionado o caixa
Public varFluxoNomeCaixa As String
Public varFluxoCodCaixa As Long
Public varFluxoCaixaSituacao As String
Public varFluxoCaixaData As String
'1.fim

Public varTipoConsulta As String
Public vOrigemRelatorio As Boolean
Public codPedido As String
Public vPedirPeso As Boolean
Public bEstNeg As Boolean       'venda de estoque negativo
Public vClienteEncontrado As Boolean

'2.impressoras
Public var_ImpNormal As String
Public var_ImpTermica As String
Public var_ImpNFCe As String
Public varImpPDF As Boolean
'2.fim

Public vChamouCaixa As String 'pdv

Public vCodFunc As Long              'codigo do funcionario para identificação
Public varValorEstimado As Double       'usando para quando apertar f2 ele mostrar o valor estimado em %
Public varCustoEstimado As Currency     'usando para quando apertar f2 ele mostrar o valor estimado em %

Public appPathApp As String      'Armazena o diretório do programa
Public appPathIni As String      'Armazena o local do arquio ini
Public appPathRpt As String      'Armazena o diretório de relatórios
Public appEXEName As String      'Armazena o nome do executável do programa
'Public appIDEmpresa As String    'Armazena o ID da empresa para validação da licença
'Public appLicenca As String      'Armazena a licença de uso do aplicativo
'Public appURLUpdt As String      'Armazena o local de atualização

Public oCfg As ConfigItem      'Arquivo ini
Public oIni As Ini             'Arquivo ini
Public var_IP As String        'Arquivo ini

Public FormParent As String      'Formulário de origem
Public CloseSystem As Boolean    'O sistema está sendo fechado

'Constantes utilizadas no projeto
Public Const ocPrjName = "Online Commerce"      'Nome do projeto
Public Const ocArqvINI = "oc.ini"               'Nome do arquivo de configurações
Public Const ocArqvRes = "ocres.dll"            'Nome do arquivo de recursos

'Constantes para formatação
Public Const ocKEYTAB = 9                       'Tab
Public Const ocKEYENTER = 13                    'enter
Public Const ocMONEY = "###,###,###,##0.00"     'Números
Public Const ocMONEY4 = "###,###,###,##0.0000"  'Números
Public Const ocPESO = "###,###,###,##0.000"     'Peso de produtos

Public Const ocCNPJ = "00\.000\.000/0000-00"    'CNPJ
Public Const ocCPF = "000\.000\.000-00"         'CPF
Public Const ocCEP = "00000-000"                'CEP
Public Const ocPHONE = "(00)0000-0000"          'telefone/fax
Public Const ocPLACA = "@@@-@@@@"               'placa
Public Const ocDATA = "dd/mm/yyyy"              'data
Public Const ocHORA = "hh:nn:ss"                'hora
Public Const ocHRMN = "hh:nn"                   'hora em minuto
Public Const ocDTHR = "dd/mm/yyyy hh:nn:ss"     'data e hora
Public Const ocDTHM = "dd/mm/yyyy hh:nn"

Public Const ocDATA_EUA = "yyyy-mm-dd"          'data formato americano
Public Const ocDTHR_EUA = "yyyy-mm-dd hh:nn:ss" 'data e hora formato americano
Public Const ocDTHM_EUA = "yyyy-mm-dd hh:nn"

'Global BD As Database
'Global db As DAO.Database
'Global AreaTrabalho As Workspace
'Public DBPath As String
'Public Ret As String

'Dim RetLen As String

'função para OS_Consulta
Public TIPO_STATUS As String
Public Condicao(1 To 4) As Variant

'variaveis para verificar se o program tá aberto
Private Const TH32CS_SNAPPROCESS As Long = 2
Private Const MAX_PATH As Long = 260

Private Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long
   th32DefaultHeapID As Long
   th32ModuleID As Long
   cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
   szExeFile As String * MAX_PATH
End Type

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, typProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, typProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

'som no windows
Private Declare Function Beep Lib "Kernel32.dll" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

'Declarações API
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const CNPJSoftHouse = "02.382.419/0001-80"
Public UTC As String

'Funcao que retorna se existe um valor em uma string
Public Function SelProcuraValor(pConteudo As String, pValor As String) As Boolean
  SelProcuraValor = InStr(1, pConteudo, " " & pValor & ",")
End Function

'Sub que adiciona em uma string, um valor, ou remove, caso ja existe
Public Sub SelAdicionaValor(ByRef pConteudo As String, ByVal pValor As String)
 If SelProcuraValor(pConteudo, pValor) Then
   pConteudo = Replace(pConteudo, " " & pValor & ",", "")
 Else
   pConteudo = pConteudo & " " & pValor & ","
 End If
End Sub

'Funcao que retorna string contendo as selecoes pronta para usar
Public Function SelValor(pConteudo As String) As String
  If Len(pConteudo) = 0 Then Exit Function
  SelValor = Left(pConteudo, Len(pConteudo) - 1)
End Function
Public Function EstoqueVendas(ByVal codProduto As String) As Double
   On Local Error GoTo errHandle
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim sldEstoque As Double
   Dim sldVenda As Double
   
   'Inicializa os saldos
   sldEstoque = 0:   sldVenda = 0
   
   'Consulta os saldos
   sSQL = "SELECT quant_estoque FROM produtos WHERE (codigo = " & codProduto & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then sldEstoque = r("quant_estoque")
      If r.State <> 0 Then r.Close
   Set r = Nothing
   
   sSQL = "SELECT ISNULL(SUM(pedidos_itens.quantidade), 0) AS quant_venda FROM pedidos_itens " & _
      "INNER JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
      "WHERE (pedidos_itens.cod_produto = " & codProduto & ") AND (pedidos_itens.data = CONVERT(DATETIME, '" & Format$(Date, ocDATA) & "', 103)) " & _
      "AND (pedidos.status_pedido = 0);"
   
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then sldVenda = r("quant_venda")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'Retorna o novo saldo em estoque
   EstoqueVendas = sldEstoque - sldVenda
   Exit Function
   
errHandle:
   EstoqueVendas = 0
End Function
Sub HabilitaObjetosVenda(Status As Boolean)
PDV.txtCodBarra.Enabled = Not Status
PDV.txtValor.Enabled = Not Status
PDV.txtQuant.Enabled = Not Status
PDV.txtTotal.Enabled = Not Status
PDV.txtTotalGeral.Enabled = Not Status
End Sub
Public Sub Main()
'Previne a execução de mais de uma vez do sistema
If App.PrevInstance Then
   ShowMsg "O sistema já encontra-se em execução nesta máquina!", vbInformation
   End
   Exit Sub
End If

ChDir App.Path                          'Muda o diretório padrão para onde está o sistema
appPathApp = App.Path                   'Armazena o diretório do sistema
NormalizePath appPathApp                'Normaliza o diretório
appPathIni = appPathApp & ocArqvINI     'Armazena o arquivo ini
appEXEName = App.EXEName & ".exe"

'Inicializa o sistema
IniciarPrograma True

'Armazena as configurações do sistema
LerConfiguracao
'Carrega o formulário de senha
'Estonar.Show
'Parcelas.Show
'Principal_Caixa.Show
'Caixa_Controle_semOS.Show
'NFCe_Consultar.Show
'Caixa_Fechamento.Show
'NFCe_Consultar.Show
PDV.Show
End Sub

'Recupera a configuração do sistema
Public Sub LerConfiguracao()
   Dim sSQL As String            'Declara as variáveis
   Dim r As ADODB.Recordset
   Dim oCfg As ConfigItem
   
   Dim vValue As Variant
   Dim lDC As Long
   Dim cIni As Ini
   
   'Lê as configurações do banco de dados
   'Essas configurações são globais
   sSQL = "SELECT config_nome, config_valor FROM configuracao ORDER BY config_nome;"
   Set r = dbData.OpenRecordset(sSQL)
   'r.Open sSQL, dbData.ActiveConnection
   
   'Inicializa a coleção de configurações globais
   Set sysConfig = Nothing
   Set sysConfig = New Collection
   
   'Percorre a tabela até o fim
   Do While Not r.EOF
      'Cria um objeto ConfigItem e atribui os valores para cada configuração
      Set oCfg = New ConfigItem
      oCfg.SetValues r("config_nome"), r("config_valor")
      sysConfig.Add oCfg, oCfg.Name
      Set oCfg = Nothing
      r.MoveNext
   Loop
   
   'Fecha a tabela
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'Inicializa a coleção de configurações locais
   Set maqConfig = Nothing
   Set maqConfig = New Collection
   
   'Inicializa o objeto de controle de arquivos INI
   Set cIni = New Ini
   
   'Seta o nome do arquivo
   cIni.Arquivo = appPathIni
   
   'Recupera a configuração de atualização
   'vValue = cIni.LerTexto("GERAL", "URLAtualizacao", "\\HI-TECH02\PUBLICA\SOFTWARE\")
   'appURLUpdt = vValue
   
   'Destrói o objeto
   Set cIni = Nothing
End Sub

'Grava a configuração do sistema
Public Sub GravarConfiguracao()
   Dim cIni As Ini
   
   'Inicializa o objeto de controle de arquivos INI
   Set cIni = New Ini
   
   'Seta o nome do arquivo
   cIni.Arquivo = appPathIni
   
   'Grava os novos valores da pasta da nfe
   'cIni.EscreverTexto "NFE", "PastaNFe", appPathNFe
   
   'Destrói o objeto
   Set cIni = Nothing
End Sub

'Mosta a MsgBox personalizada com o título do programa e retornar qual ação foi escolhida pelo usuário
Public Function ShowMsg(Prompt As String, Buttons As Integer) As Integer
   ShowMsg = MsgBox(Prompt, Buttons, ocPrjName)
End Function

'Exectua a sequencia de inicialização do programa,
'retorna True caso não ocorra erros e False no caso de alguma falha
Public Function IniciarPrograma(ExibirStatus As Boolean) As Boolean
   'Inicia o controle de erro
   On Local Error GoTo errHandle
   
   'Exibe mensagem de andamento
   'If ExibirStatus Then MsgInfo "Estabelecendo conexão ao servidor de banco de dados..."
   
   'Abre a conexão com os bancos de dados, em caso de falha
   'exibe uma mensagem de alerta e finaliza o sistema
   If Not AbrirConexaoBD Then
      ShowMsg "Não foi possível estabelecer uma conexão com o banco de dados.", vbCritical
      End
      Exit Function
   End If
   
   'Exibe mensagem de andamento
   'If ExibirStatus Then MsgInfo "Conectado ao servidor de banco de dados"
   
   'Realiza uma pausa de 1 segundo
   Sleep 1000
   
   'Exibe mensagem de andamento
   'If ExibirStatus Then MsgInfo "Verificando usuário e senha..."
   
   IniciarPrograma = True     'Retorna resultado da função
   Exit Function              'Sai da função
   
errHandle:
   'Retorna resultado de erro
   IniciarPrograma = False
End Function


Public Sub EncerrarPrograma()
'Inicia o controle de erro
On Local Error Resume Next
Dim i As Integer

'CloseThemeSupport

'Finaliza o chat
CloseSystem = True
For i = Forms.Count - 1 To 1
   Unload Forms(i)
Next

'Fecha o form principal
'Unload frmMain
   
'Verifica se as conexões foram criadas e estabelecidas,
'se True fecha todas
dbData.CloseConnection

'Finaliza todas as variáveis
Set dbData = Nothing
  
Set sysConfig = Nothing
Set maqConfig = Nothing

KillApp appEXEName
End
End Sub

Sub KillApp(appName As String)
'rotina para tirar o programa da memoria
Dim comando As String
comando = "TASKKILL -F -IM " & "PDV.exe"
Shell comando
End Sub
Public Sub KillProcess(ByVal processName As String)
Dim oWMI As Object
Dim oServices As Object
Dim oService As Object
Dim oWMIServices As Object
Dim oWMIService As Object

Dim Ret As Long
Dim sService As String
Dim servicename As String

Set oWMI = GetObject("winmgmts:")
Set oServices = oWMI.InstancesOf("win32_process")

For Each oService In oServices
    servicename = LCase(Trim(CStr(oService.Name) & ""))

    If InStr(1, servicename, LCase(processName), vbTextCompare) > 0 Then
        Ret = oService.Terminate
    End If
Next

Set oServices = Nothing
Set oWMI = Nothing
End Sub
Public Function AbrirConexaoBD() As Boolean
On Local Error GoTo errHandle   'Inicia o controle de erro
Dim cn1 As String, cn2 As String

'pegar dados no arquivo txt
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_IP = oIni.LerTexto("IP_MAQUINA", "ip")
Set oIni = Nothing
'var_IP = "192.168.1.20\SQLEXPRESS2008"
vgServerName = var_IP

'Atribui falha na execução
AbrirConexaoBD = False

'Conexão padrão do MySql
cn1 = "Provider=SQLOLEDB.1;Persist Security Info=False;DRIVER={Sql Server};SERVER=" + var_IP + ";uid=sa;pwd=190106web;DATABASE=cyber_base;TRUSTED_CONNECTION=NO"
'cn1 = "Provider=SQLOLEDB.1;Persist Security Info=False;DRIVER={Sql Server};SERVER=" + var_IP + ";uid=lotesis;pwd=lotesis;DATABASE=cyber_base;TRUSTED_CONNECTION=NO"
Set dbData = New Database

'Abre as conexões com os bancos de dados, em caso de erro sai da função
If Not dbData.OpenConnection(cn1) Then Exit Function

AbrirConexaoBD = AbreBancoDeDados    'Conexão estabelecida
Exit Function                        'Sai da função

errHandle:
   'Conexão não estabelecida
   AbrirConexaoBD = False
End Function

'Exibe uma mensagem de erro/aviso padronizada mostrando
'informações completas e descritivas do erro.
Public Sub msgErro(ByVal vModulo As String, ByVal vFuncao As String, ByVal vNumero As Long, ByVal vDescricao As String, ByVal vLinha As Integer, ByVal vOpcoes As Integer, ByVal vTipo As Integer, Optional ByVal vExibir As Boolean = True, Optional ByVal vSalvarLog As Boolean = True)
   On Local Error Resume Next
   If vExibir Then
      ShowMsg "!!! ATENÇÃO !!!" & vbNewLine & vbNewLine & _
         "Data:" & vbTab & Format$(Now, ocDATA) & vbNewLine & _
         "Hora:" & vbTab & Format$(Now, ocHORA) & vbNewLine & _
         "Projeto:" & vbTab & "Online Commerce" & vbNewLine & vbNewLine & _
         "Módulo:" & vbTab & vbTab & vModulo & vbNewLine & _
         "Procedimento:" & vbTab & vFuncao & vbNewLine & _
         "Linha:" & vbTab & vbTab & vLinha & vbNewLine & _
         "Número:" & vbTab & vbTab & vNumero & vbNewLine & _
         "Descrição:" & vbTab & vDescricao & vbNewLine, vOpcoes
   End If
   
   On Error GoTo 0
   Err.Clear
End Sub

'Calcula parcela de venda
Public Function CalculaParcela(ByVal Principal As Currency, ByVal ENTRADA As Currency, ByVal JurosAM As Currency, ByVal Parcelas As Integer) As Currency
   On Error Resume Next
   Dim cDen As Currency
   Dim cJuros As Currency
   Dim cParcela As Currency
   Dim TotalReajuste As Currency
   Dim i As Integer
   
   cDen = 1
   
   For i = 1 To Parcelas - 1
      cJuros = ((1 + (JurosAM / 100)) ^ i)
      cDen = cDen + cJuros
   Next
   
   TotalReajuste = ((Principal - ENTRADA) * ((1 + (JurosAM / 100)) ^ (Parcelas - IIf(ENTRADA <> 0, 1, 0))))
   cParcela = TotalReajuste / cDen
   CalculaParcela = Format(cParcela, "currency")
End Function

Public Sub Monta_Condicao(chkComecar As CheckBox, chkExecucao As CheckBox, chkAguardando As CheckBox, chkTerminado As CheckBox)
   Dim i As Integer
   Dim Criteria2 As String
   Dim Criteria3 As String
   Dim Criteria4 As String
   
   Condicao(1) = chkComecar.Value
   Condicao(2) = chkExecucao.Value
   Condicao(3) = chkAguardando.Value
   Condicao(4) = chkTerminado.Value

   'Limpando a variável sempre que iniciar
   TIPO_STATUS = ""
   
   For i = 1 To 4
      Select Case i
         Case 1
            If Condicao(1) = 1 Then TIPO_STATUS = TIPO_STATUS & " AND OS.STATUS = 'À COMEÇAR'"
         Case 2
            If Condicao(2) = 1 Then
               If Condicao(1) = 1 Then Criteria2 = " OR" Else Criteria2 = " AND"
               TIPO_STATUS = TIPO_STATUS & Criteria2 & " OS.STATUS = 'EM EXECUÇÃO'"
            End If
         Case 3
            If Condicao(3) = 1 Then
               If Condicao(1) = 1 Or Condicao(2) = 1 Then Criteria3 = " OR" Else Criteria3 = " AND"
               TIPO_STATUS = TIPO_STATUS & Criteria3 & " OS.STATUS = 'AGUARDANDO'"
            End If
         Case 4
            If Condicao(4) = 1 Then
               If Condicao(1) = 1 Or Condicao(2) = 1 Or Condicao(3) = 1 Then Criteria4 = " OR" Else Criteria4 = "AND"
               TIPO_STATUS = TIPO_STATUS & Criteria4 & " OS.STATUS = 'TERMINADO'"
            End If
      End Select
   Next
End Sub

Public Function aNumeros(ByVal KeyAscii As Integer, Optional Virgula As Boolean = False, Optional Ponto As Boolean = False) As Integer
Dim iRet As Integer
'Função para permitir números, vírgulas e ponto

Select Case KeyAscii
   Case 8, 13: iRet = KeyAscii
   Case 44: iRet = IIf(Virgula, 44, 0)
   Case 46: iRet = IIf(Ponto, 46, 0)
   Case 48 To 57: iRet = KeyAscii
   Case Else: iRet = 0
End Select

'Retorna a tecla pressionada
aNumeros = iRet
End Function

Public Function AppIsRunning(ByVal appName As String) As Boolean
    'rotina para verificar se um executavel está aberto
    Dim Process As PROCESSENTRY32
    Dim hSnapShot As Long
    Dim r As Long
    
    appName = LCase$(appName)
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    
    If hSnapShot <> -1 Then
        Process.dwSize = Len(Process)
        r = Process32First(hSnapShot, Process)
        Do While r
            If LCase$(Left$(Process.szExeFile, InStr(1, Process.szExeFile, vbNullChar) - 1)) = appName Then
                AppIsRunning = True
                r = False
            End If
            r = Process32Next(hSnapShot, Process)
        Loop
        CloseHandle hSnapShot
    End If
End Function
