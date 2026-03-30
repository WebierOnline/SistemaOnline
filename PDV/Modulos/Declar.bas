Attribute VB_Name = "mFuncoes"
Declare Function PegaPeso Lib "P05.DLL" (ByVal OpcaoEscrita As Long, ByVal Peso As String, ByVal Diretorio As String) As Long
Declare Function AbrePorta Lib "P05.DLL" (ByVal Porta As Long, ByVal BaudRate As Long, ByVal DataBits As Long, ByVal Paridade As Long) As Long
Declare Function FechaPorta Lib "P05.DLL" () As Long
Declare Function FechaPortaP05 Lib "P05.DLL" () As Long
Declare Sub VersaoDLL Lib "P05.DLL" (ByVal Versao As String)
Declare Function DeterminaUmStopBit Lib "P05.DLL" () As Long
