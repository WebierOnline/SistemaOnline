VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mConn As ADODB.Connection
Private mTransaction As Long
Private mConnectionString As String
Private mUltimaFalhaConexao As Boolean   'True se a ultima OpenRecordset falhou por perda de conexao (nao e ausencia real de dados)
Private mAvisoConexaoAtivo As Boolean            'True enquanto o aviso de conexao perdida ja foi mostrado e a conexao ainda nao voltou (evita repetir a cada tentativa)

Public Property Get ActiveConnection() As ADODB.Connection
   On Local Error GoTo errHandle
   Set ActiveConnection = mConn
   Exit Property
   
errHandle:
   Set ActiveConnection = Nothing
End Property

Public Property Get FalhouPorConexao() As Boolean
   FalhouPorConexao = mUltimaFalhaConexao
End Property

'Mostra o aviso de conexao perdida no maximo 1x a cada 5 segundos, evitando cascata de msgboxes
'quando varias chamadas (Verificar_Caixa, MostrarGrid_Produtos, pesquisa de produto, etc) falham em sequencia
Private Sub AvisarConexaoPerdida()
   If Not mAvisoConexaoAtivo Then
      mAvisoConexaoAtivo = True
      ShowMsg "CONEXÃO COM O SERVIDOR FOI PERDIDA!" & vbCr & _
         "Verifique se o servidor está ligado e se a rede está conectada.", vbCritical
   End If
End Sub

Public Function OpenConnection(ByVal ConnectionString As String) As Boolean
   On Error GoTo errHandle
   
   'Atribui falha na execução
   OpenConnection = False
   
   'Verifica se há uma conexão existente
   If mConn Is Nothing Then
      Set mConn = New ADODB.Connection
   End If
   
   mConnectionString = ConnectionString   'guarda para permitir reconexao automatica
   
   'Abre a conexão
   mConn.Open ConnectionString
   
   'Atribui sucesso na operação
   OpenConnection = True
   
   'Sai da função
   Exit Function
   
errHandle:
   Set mConn = Nothing   'sem popup tecnico aqui - quem chama (IniciarPrograma) mostra a mensagem amigavel unica
End Function

'Tenta reabrir a conexao usando a mesma connection string, apos queda de rede/servidor
Private Function Reconectar() As Boolean
   On Error GoTo errHandle
   Reconectar = False
   
   If mConnectionString = "" Then Exit Function
   
   Dim vConnStringRapida As String
   Dim vPosTimeout As Long, vPosFimTimeout As Long
   vConnStringRapida = mConnectionString
   
   'troca o "Connect Timeout=N" (qualquer valor de N) por um valor curto, so para a tentativa de reconexao
   'a string tem precedencia sobre a propriedade ConnectionTimeout nesse provider, nao basta trocar via propriedade
   vPosTimeout = InStr(vConnStringRapida, "Connect Timeout=")
   If vPosTimeout > 0 Then
      vPosFimTimeout = InStr(vPosTimeout, vConnStringRapida, ";")
      If vPosFimTimeout > 0 Then
         vConnStringRapida = Left(vConnStringRapida, vPosTimeout - 1) & "Connect Timeout=5" & Mid(vConnStringRapida, vPosFimTimeout)
      End If
   End If
   
   Set mConn = New ADODB.Connection
   mConn.ConnectionTimeout = 5   'timeout curto so para a tentativa de reconexao, nao trava a tela por 600s
   mConn.Open vConnStringRapida
   Reconectar = True
   Exit Function
   
errHandle:
   Set mConn = Nothing
   Reconectar = False
End Function

Public Sub CloseConnection()
   On Error GoTo errHandle
   
   'Verifica se há uma conexão existente
   If Not mConn Is Nothing Then
      'Verifica se a conexão está aberta
      If mConn.State <> 0 Then
         'Verifica se há alguma transação pendente e desfaz
         If mTransaction <> 0 Then mConn.RollbackTrans
         'Fecha a conexão
         mConn.Close
      End If
   End If
   
   'Sai da rotina
   Exit Sub
   
errHandle:
   Set mConn = Nothing
   msgErro "Database", "CloseConnection", Err.Number, Err.Description, Erl, vbCritical, 2
End Sub

'Esta função somente funciona perfeitamente com o SQL Server, evite usar em outros servidores com o MySQL
Public Function OpenRecordset(ByVal SQLCommand As String, Optional ByRef RecordCount As Long, Optional ByVal vTentativaReconexao As Integer = 0) As ADODB.Recordset
   On Error GoTo errHandle
   Dim rs As ADODB.Recordset
   mUltimaFalhaConexao = False
   
   'Inicializa as variáveis
   Set rs = New ADODB.Recordset
   'Cursor no lado do cliente
   rs.CursorLocation = adUseClient
   rs.CursorType = adOpenStatic
   rs.LockType = adLockReadOnly
   
   'Tenta reconectar automaticamente se a conexao caiu
   If mConn Is Nothing Then
      Reconectar
   ElseIf mConn.State = 0 Then
      Reconectar
   End If
   
   'Se a reconexao falhou, mConn continua Nothing - nao tenta abrir (evita erro 91)
   If mConn Is Nothing Then
      mUltimaFalhaConexao = True
      AvisarConexaoPerdida
      GoSub errReturnEmpty
      Exit Function
   End If
   
   'Abre o banco de dados selecionado
   rs.Open SQLCommand, mConn
   mAvisoConexaoAtivo = False   'conexao confirmada funcionando, permite avisar novamente numa proxima queda
   
   'Armazena a contagem de registro
   If Not rs.BOF Then RecordCount = rs.RecordCount
   
   'Desconecta do servidor
   Set rs.ActiveConnection = Nothing
   
   'Retorna o resultado da consulta
   Set OpenRecordset = rs.Clone
   
   'Fecha a tabela e destrói a variável da memória
   If rs.State <> 0 Then rs.Close
   Set rs = Nothing
   
   'Sai da função
   Exit Function
   
errReturnEmpty:
   RecordCount = 0
   Set OpenRecordset = New ADODB.Recordset
   OpenRecordset.CursorLocation = adUseClient
   OpenRecordset.Fields.Append "codigo", adInteger
   OpenRecordset.Open  'garante recordset aberto (EOF/BOF=True), evita erro 3704 no chamador
   Set OpenRecordset.ActiveConnection = Nothing
   Return
   
errHandle:
   Select Case Err.Number
      Case -2147467259        'Falha de conexão (rede/servidor indisponível)
         Set mConn = Nothing   'forca nova tentativa de reconexao na proxima chamada
         If vTentativaReconexao = 0 And Reconectar() Then
            Set OpenRecordset = OpenRecordset(SQLCommand, RecordCount, 1)
            Exit Function
         End If
         mUltimaFalhaConexao = True
         AvisarConexaoPerdida
         GoSub errReturnEmpty
      
      Case -2147217871        'Tempo excedido
         ShowMsg "Tempo limite para execução da operação foi expirado." & vbCr & _
            "Tente novamente, e em caso do erro persistir contacte o suporte técnico.", vbCritical
         GoSub errReturnEmpty
         
      Case -2147217900        'Sintaxe SQL incorreta
         ShowMsg "Comando inválido." & vbCr & vbCr & _
            "Descrição: " & Err.Description & vbCr & vbCr & _
            "Linha: " & SQLCommand, vbCritical
         GoSub errReturnEmpty
      
      Case Else
         If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
         Set rs = Nothing
         Set OpenRecordset = Nothing
         msgErro "Database", "OpenRecordset", Err.Number, Err.Description, Erl, vbCritical, 2
   End Select
End Function

Public Function BeginTransaction() As Boolean
   On Error GoTo errHandle
   
   BeginTransaction = False            'Atribui falha na execução
   'Nenhum transação foi encontrada, inicia uma nova
   If mTransaction = 0 Then mTransaction = mConn.BeginTrans
   BeginTransaction = True             'Atribui sucesso na execução
   Exit Function                       'Sai da função
   
errHandle:
   BeginTransaction = False            'Atribui falha
   
   'Se há uma conexão existente
   If Not mConn Is Nothing Then
      If mTransaction <> 0 Then mConn.RollbackTrans   'Desfaz as alterações
   End If
   
   mTransaction = 0
   msgErro "Classe Database", "BeginTransaction", Err.Number, Err.Description, Erl, vbCritical, 2
End Function

Public Function CommitTransaction(Optional ByVal KeepAlive As Boolean = False) As Boolean
   On Error GoTo errHandle
   
   'Atribui falha na execução
   CommitTransaction = False
   
   'Verifica se há uma conexão existente
   If Not mConn Is Nothing Then
      If mTransaction <> 0 Then  'Verifica se há uma transação aberta
         'Não mantém a transação
         If (Not KeepAlive) Then
            mConn.CommitTrans       'Grava as alterações
            mTransaction = 0        'Limpa a transação
         End If
      End If
   End If
   
   CommitTransaction = True      'Atribui sucesso na execução
   Exit Function                 'Sai da função
   
errHandle:
   RollbackTransaction
   msgErro "Classe Database", "CommitTransaction", Err.Number, Err.Description, Erl, vbCritical, 2
End Function

Public Sub RollbackTransaction()
   On Error GoTo errHandle
   
   'Verifica se há uma conexão existente
   If Not mConn Is Nothing Then
      If mTransaction <> 0 Then     'Verifica se há uma transação aberta
         mConn.RollbackTrans        'Desfaz as alterações
      End If
   End If
      
   mTransaction = 0        'Limpa a transação
   Exit Sub                'Sai da rotina
   
errHandle:
   'Se há uma conexão existente
   If Not mConn Is Nothing Then
      If mTransaction <> 0 Then mConn.RollbackTrans   'Desfaz as alterações
   End If
   
   mTransaction = 0
   msgErro "Classe Database", "RollbackTransaction", Err.Number, Err.Description, Erl, vbCritical, 2
End Sub

Public Function Execute(ByVal CommandText As String, Optional ByRef RecordsAffected As Long) As Boolean
On Error GoTo errHandle
Dim lRec As Long

'Atribui falha na execução
Execute = False
'Inicializa o contador de registros
lRec = 0

'Tenta reconectar automaticamente se a conexao caiu
If mConn Is Nothing Then
   Reconectar
ElseIf mConn.State = 0 Then
   Reconectar
End If

'Verifica se há uma conexão global
If mConn Is Nothing Then
   msgErro "Database", "Execute", 0, "Conexao com o banco de dados nao esta disponivel.", Erl, vbCritical, 2
   Exit Function
End If

mConn.Execute CommandText, lRec  'Executa o comando

'Armazena os valores das propriedades
RecordsAffected = lRec

'Retorna o resultado da função
Execute = True
Exit Function

errHandle:
Debug.Print Err.Number & vbTab & Err.Description   'desativei aqui cleia
'Se há uma conexão existente
If Not mConn Is Nothing Then
   RollbackTransaction           'Desfaz as alterações
End If

If Err.Number = -2147467259 Then Set mConn = Nothing   'forca reconexao na proxima chamada

msgErro "Database", "Execute", Err.Number, CommandText, Erl, vbCritical, 2
End Function

Public Function GetDateTime() As Date
   
   '
   
End Function

Private Sub Class_Initialize()
   'Instancia a variável de conexão
   Set mConn = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
   CloseConnection      'Fecha a conexão
   Set mConn = Nothing  'Destrói a variável
End Sub
