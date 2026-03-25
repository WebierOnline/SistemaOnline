Attribute VB_Name = "Util"
Option Explicit
Public xWinRar As String, xWinZip As String
Public vgDb As New ADODB.Connection                 'variável objeto banco de dados
Public vgCat As New ADOX.Catalog
Public vgServerName As String                       'Nome do servidor SQL

'usadas na montagem de "queries" para extrair ou colocar as cláusulas
'da expressão SQL
Public Const EXP_SELECT = 0                         'extrai a expressão SELECT/UPDATE/DELETE do SQL
Public Const EXP_SET = 1                            'extrai a expressão SET do SQL
Public Const EXP_FROM = 2                           'extrai a expressão FROM do SQL
Public Const EXP_LEFT_JOIN = 3                      'extrai a expressão LEFT JOIN do SQL
Public Const EXP_RIGHT_JOIN = 4                     'extrai a expressão RIGHT JOIN do SQL
Public Const EXP_INNER_JOIN = 5                     'extrai a expressão INNER JOIN do SQL
Public Const EXP_INNER_ON = 6                       'extrai a expressão ON do SQL
Public Const EXP_WHERE = 7                          'extrai a expressão WHERE do SQL
Public Const EXP_GROUPBY = 8                        'extrai a expressão GROUPBY do SQL
Public Const EXP_HAVING = 9                         'extrai a expressão HAVING do SQL
Public Const EXP_ORDERBY = 10                       'extrai a expressão ORDERBY do SQL
Public Const EXP_LIMIT = 11                         'extrai a expressão LIMIT do SQL (MySQL)
Public Const EXP_TODAS = 12                         'extrai a expressão SQL inteira

Public vgClausula(EXP_TODAS - 1) As String        'vetor com os nomes das cláusulas SQL

'parâmetros da função HaNaString
Public Const UM_A_UM = -1                         'só um caracter testado
Public Const SO_UM = 0                            'todos os caracteres testados um a um

' Define o tipo para área de um retângulo
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'Define o tipo para coordenadas de um ponto
Public Type POINT
   x As Long
   y As Long
End Type

Private Type TIME_OF_DAY         'Hora
   t_elapsedt As Long
   t_msecs As Long
   t_hours As Long
   t_mins As Long
   t_secs As Long
   t_hunds As Long
   t_timezone As Long
   t_tinterval As Long
   t_day As Long
   t_month As Long
   t_year As Long
   t_weekday As Long
End Type

Public Type BITMAPINFOHEADER
   biSize          As Long
   biWidth         As Long
   biHeight        As Long
   biPlanes        As Integer
   biBitCount      As Integer
   biCompression   As Long
   biSizeImage     As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed       As Long
   biClrImportant  As Long
End Type

Public Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Public Type BITMAPINFO
   bmiHeader As BITMAPINFOHEADER
   bmiColors As RGBQUAD
End Type

Private Type OSVERSIONINFO       'Versão do sistema operacional
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type

Public Type MOUSEINPUT           'Mouse
   dx As Long
   dy As Long
   mouseData As Long
   dwFlags As Long
   Time As Long
   dwExtraInfo As Long
End Type

Public Type KEYBDINPUT           'Teclado
   wVk As Integer
   wScan As Integer
   dwFlags As Long
   Time As Long
   dwExtraInfo As Long
End Type

Public Type HARDWAREINPUT        'Hardware
   uMsg As Long
   wParamL As Integer
   wParamH As Integer
End Type

Public Type GENERALINPUT         '
   dwType As Long
   xi(0 To 23) As Byte
End Type

'Constantes gerais
Public Const VK_ENTER = 13
Public Const KEYEVENTF_KEYUP = &H2
Public Const INPUT_MOUSE = 0
Public Const INPUT_KEYBOARD = 1
Public Const INPUT_HARDWARE = 2

Public Const TIME_ONESHOT = 0                   'Event occurs once, after uDelay milliseconds.
Public Const TIME_PERIODIC = 1                  'Event occurs every uDelay milliseconds.
Public Const TIME_CALLBACK_EVENT_PULSE = &H20   'When the timer expires, Windows calls thePulseEvent function to pulse the event pointed to by the lpTimeProc parameter. The dwUser parameter is ignored.
Public Const TIME_CALLBACK_EVENT_SET = &H10     'When the timer expires, Windows calls theSetEvent function to set the event pointed to by the lpTimeProc parameter. The dwUser parameter is ignored.
Public Const TIME_CALLBACK_FUNCTION = &H0       'When the timer expires, Windows calls the function pointed to by the lpTimeProc parameter. This is the default.

Public Const DIB_RGB_COLORS As Long = 0
Public Const BI_RGB = 0&

'Variáveis para verificação do agendamento do backup
Public VBTimer As Long, MMTimer As Long
Public hMMTimer As Long
Public updTimer As Long

'Declarações API
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
Public Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long

Public Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

Private Declare Function NetRemoteTOD Lib "NETAPI32.DLL" (ByVal Server As String, Buffer As Any) As Long
Private Declare Function NetApiBufferFree Lib "NETAPI32.DLL" (Buffer As Any) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpOSInfo As OSVERSIONINFO) As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)

Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long

Private Const CS_DROPSHADOW As Long = &H20000
Private Const GCL_STYLE As Long = -26

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Public Function GoogleEnviarArquivo(nomeArquivo As String) As Boolean
Dim uploader As Object
    On Error GoTo deuErro
    Set uploader = CreateObject("GoogleDriveUploader.Uploader")
    If uploader Is Nothing Then Exit Function
    uploader.UseInternalDialog = True
    uploader.ApplicationName = "OnlineInfo"
    uploader.CredentialsPath = App.Path & "\clientGoogle_secret.json"
    uploader.AuthMode = "OAuth"
    uploader.TokenPath = App.Path & "\tokens"
    uploader.UserToImpersonate = "financeiroonlineinfo@gmail.com"  ' Meu Drive deste usuário
    Dim fid As String
    fid = uploader.UploadFileToFolderName(nomeArquivo, "BACKUP")
    'txtResult.Text = fid
    GoogleEnviarArquivo = True
    Exit Function
deuErro:
    'MsgBox "Erro: " & Err.Description, vbCritical
    mensagemErro = Err.Description
    Err.Clear
End Function

'Função para absorver directorias com espaços
Public Function Transforma(Ficheiro As String) As String
    Transforma = IIf(InStr(Ficheiro, " "), """" & Ficheiro & """", Ficheiro)
End Function
'Executa o comando SQL no banco de dados
Public Function SQLExecuta(ComandoSQL As String, Optional ByRef NRegs As Long) As String
 On Error Resume Next
 'Executa o camando SQL
 dbData.Execute ComandoSQL, NRegs
 SQLExecuta = Err.Description
End Function

Public Function IniciaComponenteCompactacao() As Boolean
    On Error Resume Next
    'Verifica se está instalado um dos 2 compressores referênciados
    
    Dim MiObjeto As Object, Ret As String
    Set MiObjeto = CreateObject("Wscript.Shell")
    
    Ret = MiObjeto.regread("HKEY_CLASSES_ROOT\WinRAR\shell\open\command\")
    xWinRar = Left(Ret, InStrRev(Ret, " "))
    
    Ret = MiObjeto.regread("HKEY_CLASSES_ROOT\WinZip\shell\open\command\")
    xWinZip = Left(Ret, InStrRev(Ret, " "))
    
    IniciaComponenteCompactacao = True
    Set MiObjeto = Nothing
    
    If xWinRar & xWinZip = "" Then
       MsgBox "Não se encontra instalado o WinZip nem WinRar :(", vbCritical
       IniciaComponenteCompactacao = False
    End If
End Function
'*******************************************************************
' Procedimento: TimerProc
' Objetivo    : Fechar a MsgBox do backup em tempo programado
'*******************************************************************
Public Sub TimerProc(ByVal uID As Long, ByVal uMsg As Long, ByVal dwUser As Long, ByVal dw1 As Long, ByVal dw2 As Long)
   SendKey VK_ENTER
   timeKillEvent hMMTimer
   timeKillEvent updTimer
End Sub

Public Sub SendKey(bKey As Byte)
   Dim GInput(0 To 1) As GENERALINPUT
   Dim KInput As KEYBDINPUT
   
   KInput.wVk = bKey  'Tecla que será pressionada
   KInput.dwFlags = 0 'Pressionar a tecla
   
   'Entrada do teclado
   GInput(0).dwType = INPUT_KEYBOARD
   'Copiar a estrutura para a matriz de memória
   CopyMemory GInput(0).xi(0), KInput, Len(KInput)
   
   'Mesmo que acima, mas soltando a tecla
   KInput.wVk = bKey                 'Tecla que será solta
   KInput.dwFlags = KEYEVENTF_KEYUP  'Soltar a tecla
   GInput(1).dwType = INPUT_KEYBOARD 'Entrada do teclado
   
   'Copiar a estrutura para a matriz de memória
   CopyMemory GInput(1).xi(0), KInput, Len(KInput)
   
   'Envia a entrada agora
   Call SendInput(2, GInput(0), Len(GInput(0)))


End Sub

Public Sub ApplyDropShadow(ByVal hwnd As Long)
   SetClassLong hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub

Public Function Validar_CNPJ(ByVal CNPJ As String) As Boolean
Dim bolRetorno As Boolean
bolRetorno = False
    If Len(CNPJ) > 0 Then
        'Retiramos possíveis máscaras
        CNPJ = RetirarMascaras(CNPJ)
        '
        If Len(CNPJ) = 14 Then
            If IsNumeric(CNPJ) Then
                Dim strNumeros As String, strMultiplicador As String
                Dim intPosiçao As Integer, intDV1 As Integer, intDV2 As Integer, intResto As Integer
                '
                strNumeros = Left(CNPJ, 12)
                strMultiplicador = "543298765432"
                intDV1 = 0
                intDV2 = 0
                intPosiçao = 12
                '
                While intPosiçao > 0
                    intDV1 = intDV1 + (Val(Mid(strNumeros, intPosiçao, 1)) * Val(Mid(strMultiplicador, intPosiçao, 1)))
                    intPosiçao = intPosiçao - 1
                Wend
                '
                intResto = intDV1 Mod 11
                    If intResto < 2 Then
                    intDV1 = 0
                Else
                    intDV1 = 11 - intResto
                End If
                '
                strNumeros = strNumeros & Right(CStr(intDV1), 1)
                strMultiplicador = "6" & strMultiplicador
                intPosiçao = 13
                '
                While intPosiçao > 0
                    intDV2 = intDV2 + (Val(Mid(strNumeros, intPosiçao, 1)) * Val(Mid(strMultiplicador, intPosiçao, 1)))
                    intPosiçao = intPosiçao - 1
                Wend
                '
                intResto = intDV2 Mod 11
                    If intResto < 2 Then
                    intDV2 = 0
                Else
                    intDV2 = 11 - intResto
                End If
                '
                bolRetorno = ((intDV1 = Val(Mid(CNPJ, 13, 1))) And (intDV2 = Val(Right(CNPJ, 1))))
            End If
        End If
    End If
'---Retornar
    Validar_CNPJ = bolRetorno

End Function

Public Function RetirarMascaras(ByVal Texto As String) As String
    Texto = Replace(Texto, ".", Empty)
    Texto = Replace(Texto, "/", Empty)
    Texto = Replace(Texto, "-", Empty)
    Texto = Replace(Texto, ",", Empty)
    Texto = Replace(Texto, ":", Empty)
    Texto = Replace(Texto, "_", Empty)
    Texto = Replace(Texto, "(", Empty)
    Texto = Replace(Texto, ")", Empty)
    Texto = Replace(Texto, " ", Empty)
'---Retornar | Me.Text2.Text = RetirarMascaras(Me.Text1.Text)
    RetirarMascaras = Texto
End Function


Public Function Validar_CPF(ByVal CPF As String) As Boolean
Dim bolRetorno As Boolean
bolRetorno = False
    If Len(CPF) > 0 Then
        'Retiramos possíveis máscaras
        CPF = RetirarMascaras(CPF)
        '
        If Len(CPF) = 11 Then
            If IsNumeric(CPF) Then
                Select Case CPF
                    Case "11111111111", "22222222222", "33333333333", "44444444444", "55555555555", "66666666666", "77777777777", "88888888888", "99999999999", "00000000000"
                      'Não verificamos
                      '
                    Case Else
                        Dim strNumeros As String, strDV As String
                        Dim intDV1 As Integer, intDV2 As Integer, intSoma As Integer, intPosiçao As Integer, intResto As Integer
                        '
                        strNumeros = Left(CPF, 9)
                        intSoma = 0
                        intPosiçao = 0
                        '
                        While intPosiçao < 9
                            intPosiçao = intPosiçao + 1
                            intSoma = intSoma + (Val(Mid(strNumeros, 10 - intPosiçao, 1)) * (intPosiçao + 1))
                        Wend
                        '
                        intResto = intSoma Mod 11
                        If intResto < 2 Then
                            intDV1 = 0
                        Else
                            intDV1 = 11 - intResto
                        End If
                        '
                        strDV = Right(CStr(intDV1), 1)
                        strNumeros = strNumeros & strDV
                        intSoma = 0
                        intPosiçao = 0
                        '
                        While intPosiçao < 10
                            intPosiçao = intPosiçao + 1
                            intSoma = intSoma + (Val(Mid(strNumeros, 11 - intPosiçao, 1)) * (intPosiçao + 1))
                        Wend
                        '
                        intResto = intSoma Mod 11
                        If intResto < 2 Then
                            intDV2 = 0
                        Else
                            intDV2 = 11 - intResto
                        End If
                        strDV = strDV & Right(CStr(intDV2), 1)
                        bolRetorno = (strDV = Right(CPF, 2))
                    End Select
            End If
        End If
    End If
'---Retornar
    Validar_CPF = bolRetorno

End Function


Public Sub ApplyTransparency(ByVal hwnd As Long, ByVal lOpacity As Long)
   SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
   SetLayeredWindowAttributes hwnd, 0, lOpacity, &H2&
End Sub

Public Sub PaintGradient(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long)
  Dim uBIH    As BITMAPINFOHEADER
  Dim lBits() As Long
  Dim lGrad() As Long
  
  Dim r1 As Long, g1 As Long, b1 As Long
  Dim r2 As Long, g2 As Long, b2 As Long
  Dim dR As Long, dG As Long, dB As Long
  
  Dim Scan As Long
  Dim i As Long
  Dim iEnd    As Long
  Dim iOffset As Long
  Dim j       As Long
  Dim jEnd    As Long
  Dim iGrad   As Long
  
   '-- A minor check
   If (Width < 1 Or Height < 1) Then Exit Sub
    
   '-- Decompose colors
   Color1 = Color1 And &HFFFFFF
   r1 = Color1 Mod &H100&
   Color1 = Color1 \ &H100&
   g1 = Color1 Mod &H100&
   Color1 = Color1 \ &H100&
   b1 = Color1 Mod &H100&
   
   Color2 = Color2 And &HFFFFFF
   r2 = Color2 Mod &H100&
   Color2 = Color2 \ &H100&
   g2 = Color2 Mod &H100&
   Color2 = Color2 \ &H100&
   b2 = Color2 Mod &H100&
   
   '-- Get color distances
   dR = r2 - r1
   dG = g2 - g1
   dB = b2 - b1
   
   '-- Size gradient-colors array
   ReDim lGrad(0 To Height - 1)
   
   '-- Calculate gradient-colors
   iEnd = UBound(lGrad())
   If (iEnd = 0) Then
      '-- Special case (1-pixel wide gradient)
      lGrad(0) = (b1 \ 2 + b2 \ 2) + 256 * (g1 \ 2 + g2 \ 2) + 65536 * (r1 \ 2 + r2 \ 2)
   Else
      For i = 0 To iEnd
         lGrad(i) = b1 + (dB * i) \ iEnd + 256 * (g1 + (dG * i) \ iEnd) + 65536 * (r1 + (dR * i) \ iEnd)
      Next
   End If
   
   '-- Size DIB array
   ReDim lBits(Width * Height - 1) As Long
   iEnd = Width - 1
   jEnd = Height - 1
   Scan = Width
   
   '-- Render gradient DIB
   For j = jEnd To 0 Step -1
      For i = iOffset To iEnd + iOffset
         lBits(i) = lGrad(j)
      Next
      iOffset = iOffset + Scan
   Next
   
   '-- Define DIB header
   With uBIH
      .biSize = 40
      .biPlanes = 1
      .biBitCount = 32
      .biWidth = Width
      .biHeight = Height
   End With
   
   '-- Paint it!
   Call StretchDIBits(hDC, x, y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
End Sub

'*******************************************************************
' Procedimento: VersaoPrograma
' Argumentos  : Nenhum
' Retorno     : String
' Objetivo    : Retorna a versão do programa
'*******************************************************************
Public Function VersaoPrograma() As String
   Dim fArqv As String
   Dim rC As Long, lDummy As Long, sBuffer() As Byte
   Dim lBufferLen As Long, lVerPointer As Long, tVerBuffer As VS_FIXEDFILEINFO
   Dim lVerbufferLen As Long
   
   'Armazena o nome do programa
   fArqv = appPathApp & appEXEName
   
   '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(fArqv, lDummy)
   If lBufferLen < 1 Then
      VersaoPrograma = ""
      Exit Function
   End If
   
   '**** Store info to udtVerBuffer struct ****
   ReDim sBuffer(lBufferLen)
   rC = GetFileVersionInfo(fArqv, 0&, lBufferLen, sBuffer(0))
   rC = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
   MoveMemory tVerBuffer, lVerPointer, Len(tVerBuffer)
   
   '**** Determine File Version number ****
   VersaoPrograma = Format$(tVerBuffer.dwFileVersionMSh) & "." & Format$(tVerBuffer.dwFileVersionMSl) & "." & Format$(tVerBuffer.dwFileVersionLSh) & "." & Format$(tVerBuffer.dwFileVersionLSl)
End Function

Public Function RemoveAcento(sString As String) As String

    Dim sRet As String

    sRet = sString

    sRet = Replace(sRet, "<", " ")
    sRet = Replace(sRet, ">", " ")
    sRet = Replace(sRet, "&", "E")
    sRet = Replace(sRet, "'", " ")

    sRet = Replace(sRet, "á", "a")
    sRet = Replace(sRet, "à", "a")
    sRet = Replace(sRet, "â", "a")
    sRet = Replace(sRet, "ã", "a")
    sRet = Replace(sRet, "ä", "a")

    sRet = Replace(sRet, "é", "e")
    sRet = Replace(sRet, "è", "e")
    sRet = Replace(sRet, "ê", "e")
    sRet = Replace(sRet, "ë", "e")

    sRet = Replace(sRet, "í", "i")
    sRet = Replace(sRet, "ì", "i")
    sRet = Replace(sRet, "î", "i")
    sRet = Replace(sRet, "ï", "i")

    sRet = Replace(sRet, "ó", "o")
    sRet = Replace(sRet, "ò", "o")
    sRet = Replace(sRet, "ô", "o")
    sRet = Replace(sRet, "õ", "o")
    sRet = Replace(sRet, "ö", "o")

    sRet = Replace(sRet, "ú", "u")
    sRet = Replace(sRet, "ù", "u")
    sRet = Replace(sRet, "û", "u")
    sRet = Replace(sRet, "ü", "u")

    sRet = Replace(sRet, "ç", "c")

    sRet = Replace(sRet, "Á", "A")
    sRet = Replace(sRet, "À", "A")
    sRet = Replace(sRet, "Â", "A")
    sRet = Replace(sRet, "Ã", "A")
    sRet = Replace(sRet, "Ä", "A")

    sRet = Replace(sRet, "É", "E")
    sRet = Replace(sRet, "È", "E")
    sRet = Replace(sRet, "Ê", "E")
    sRet = Replace(sRet, "Ë", "E")

    sRet = Replace(sRet, "Í", "I")
    sRet = Replace(sRet, "Ì", "I")
    sRet = Replace(sRet, "Î", "I")
    sRet = Replace(sRet, "Ï", "I")

    sRet = Replace(sRet, "Ó", "O")
    sRet = Replace(sRet, "Ò", "O")
    sRet = Replace(sRet, "Ô", "O")
    sRet = Replace(sRet, "Õ", "O")
    sRet = Replace(sRet, "Ö", "O")

    sRet = Replace(sRet, "Ú", "U")
    sRet = Replace(sRet, "Ù", "U")
    sRet = Replace(sRet, "Û", "U")
    sRet = Replace(sRet, "Ü", "U")

    sRet = Replace(sRet, "Ç", "C")

    sRet = Replace(sRet, "°", ".")
    sRet = Replace(sRet, "º", ".")
    sRet = Replace(sRet, "ª", ".")
    
    sRet = Replace(sRet, Chr(13), " ")
    sRet = Replace(sRet, Chr(10), " ")
    sRet = Replace(sRet, vbNewLine, " ")
    sRet = Replace(sRet, "  ", " ")
    
    sRet = Replace(sRet, "§", "INCISO(S)")
    
    sRet = LTrim(sRet)
    sRet = RTrim(sRet)

    'RemoveAcento = UCase(sRet)
    RemoveAcento = sRet


End Function

'*******************************************************************
' Procedimento: CenterForm
' Argumentos  : FormObj As Form
'               -> Form a ser centralizado
' Retorno     : Nenhum
' Objetivo    : Centraliza o form na tela
'*******************************************************************
Public Sub CenterForm(FormObj As Form, Optional ByVal NewWidth As Long, Optional NewHeight As Long)
   Dim lLft As Long, lTop As Long
   Dim lWdt As Long, lHgt As Long
   
   'Cálcula a posição
   lLft = (Screen.Width - FormObj.ScaleWidth) / 2
   lTop = (Screen.Height - FormObj.ScaleHeight) / 2
   
   'Cálcula o novo tamanho
   lWdt = FormObj.Width
   lHgt = FormObj.Height
   If NewWidth <> lWdt Then lWdt = NewWidth
   If NewHeight <> lHgt Then lHgt = NewHeight
   
   'Centraliza o form
   FormObj.Move lLft, lTop, lWdt, lHgt
End Sub

Public Function Truncate(ByVal Numero As Double, ByVal Fator As Byte) As Double
   Truncate = Fix(Numero * 10 ^ Fator) / 10 ^ Fator
End Function

'*******************************************************************
' Procedimento: ArredondarMoeda
' Argumentos  : Valor As Currency
'               -> Valor a ser arredondado
' Retorno     : String
' Objetivo    : Arredonda um valor monetário em 2 casas decimais
'*******************************************************************
Public Function ArredondarMoeda(ByVal Valor As Currency) As String
   Dim iPos As Integer     'Declara as variáveis
   Dim sMoeda As String
   
   'Converte para string
   sMoeda = Valor
   'Verifica qual a posição da vírgula dentro da variável
   iPos = InStr(1, sMoeda, ",")
   'Se não há vígula o número é inteiro, então a posição é o tamanho da string
   If iPos = 0 Then iPos = Len(sMoeda)
   'Retorna o número até a segunda casa decimal
   ArredondarMoeda = Mid$(sMoeda, 1, iPos + 2)
End Function

Public Sub CalcularParcelas(ByVal ValorTotal As Currency, ByVal NroParcelas As Integer, ByRef ValorParcelas() As Currency)
   'Declara as variáveis
   Dim i As Integer, j As Integer
   
   Dim vlr As Currency, svlr As String
   Dim dif As Long, totParc As Currency

   'Cria a lista de parcelas
   ReDim ValorParcelas(1 To NroParcelas)
   
   'Calcula o valor de cada parcela
   vlr = ValorTotal / NroParcelas
   'Recupera somente as 2 casas decimais após a vírgula
   vlr = ArredondarMoeda(vlr)
   
   'Percorre a lista de parcelas da última para a primeira
   For i = 1 To NroParcelas
      ValorParcelas(i) = vlr        'Armazena o valor da parcela
      totParc = totParc + vlr       'Totaliza o sub total das parcelas
   Next
   
   'Verifica a diferença entre as parcelas e o valor total
   dif = Format$(Abs(ValorTotal - totParc) * 100, ocMONEY)
   
   'Este cálculo permite redividir as parcelas quando a
   'divisão não é exata e retorna alguns centavos de diferença
   'esta rotina, soma ou deduz R$ 0,01 para ajustar o valor
   'das parcelas para que a soma delas seja igual ao valor
   'do subtotal da filial
   
   'Executa o número de vezes da diferença obtida
   For i = 1 To dif
      'Verifica se a soma das parcelas é maior que o subtotal
      'Neste caso, deduz R$ 0,01 de cada parcela
      If totParc > ValorTotal Then
         ValorParcelas(i) = ValorParcelas(i) - 0.01
         
      'A soma das parcels é menor que o subtotal
      'Neste caso, soma R$ 0,01 para cada parcela
      Else
         ValorParcelas(i) = ValorParcelas(i) + 0.01
      
      End If
   Next
End Sub

Public Function ArredondarPBaixo(ByVal Valor As Currency) As Currency
   Dim iPos As Integer
   Dim sMoeda As String, vMoeda As Double
   Dim iMoeda As Long
   Dim vUni As Integer
   
   'Converte para string
   sMoeda = CStr(Valor)
   'Verifica qual a posição da vírgula dentro da variável
   iPos = InStr(1, sMoeda, ",")
   'Se não há vígula o número é inteiro, então a posição é o tamanho da string
   If iPos = 0 Then iPos = Len(sMoeda)
   'Retorna o número até a segunda casa decimal
   vMoeda = Mid(sMoeda, 1, iPos + 2)
   
   'Retorna somente o valor inteiro
   iMoeda = Int(vMoeda)
   
   'Seleciona a unidade do valor
   vUni = Right$(iMoeda, 1)
   
   'Escolhe a opção
   Select Case vUni
      Case 0, 5: vMoeda = iMoeda
      Case 1 To 4: vMoeda = iMoeda - vUni
      Case 6 To 9: vMoeda = iMoeda - (vUni - 5)
   End Select
   
   'Retorna o arredondamento
   ArredondarPBaixo = vMoeda
End Function

'*******************************************************************
' Procedimento: ConverterHoraExtenso
' Argumentos  : TimeMin As Long
'               -> Tempo em minutos
'               ShowHourWhenZero As Boolean (Opcional)
'               -> Exibe ou não a hora quando esta for 0.
' Retorno     : String
' Objetivo    : Converte uma quantidade de tempo em minutos para o
'               formato de hora (00:00)
'*******************************************************************
Public Function ConverterHoraExtenso(ByVal TimeMin As Long, Optional ByVal ExibirHoraQuandoZero As Boolean = True) As String
   Dim hrs As Long, min As Long  'Declara as variáveis
   Dim sTime As String
   
   sTime = ""                 'Inicializa a variável
   hrs = Abs(TimeMin) \ 60    'Armazena a qtde de horas
   min = Abs(TimeMin) Mod 60  'Armazena a qtde de minutos
   
   'If hrs > 0 Then sTime = hrs & " h e "
   'If hrs > 0 And min > 0 Then sTime = sTime & ""
   'If min > 0 Then sTime = sTime & min & " min"
   
   'Cria a conversão por extenso
   If hrs = 0 Then
      If ExibirHoraQuandoZero Then sTime = hrs & " h e "
   Else
      sTime = hrs & " h e "
   End If
   
   'Completa a hora com os minutos
   sTime = sTime & Format$(min, "00") & " min"
   
   'Retorna o resultado
   ConverterHoraExtenso = sTime
End Function

Public Function ConverterHoraParaMinuto(ByVal Horas As Long) As Long
   Dim min As Long      'Declara as variáveis
   min = Horas * 60                    'Calcula os minutos
   ConverterHoraParaMinuto = Fix(min)  'Retorna o resultado
End Function

'*******************************************************************
' Procedimento: MaskEditNumber
' Argumentos  : Ctl As TextoBox
'               -> Qualquer caixa de texto para formatação
'               FloatPoint As Integer
'               -> Define o número de casas decimais
'               KeyAscii As Integer
'               -> Recebe a tecla pressiona pelo usuário
' Retorno     : Nenhum
' Objetivo    : Criar uma máscara para digitação de números nas
'               caixas de texto
'*******************************************************************
Public Sub MaskEditNumber(ByVal ctl As Control, ByVal FloatPoint As Integer, KeyAscii As Integer)
   Dim hasPoint As Integer
   Dim hasTwoNumber As Integer
   
   With ctl
      Select Case KeyAscii
         Case 8: If Not .Locked Then .SelText = ""
         Case 13: SendKey ocKEYTAB
         Case 44, 45, 46, 48 To 57
            If .Locked Then Exit Sub
            
            If Not .SelLength > 0 Then
               If InStr(1, .Text, Chr$(44)) > 0 Then
                  hasTwoNumber = Len(Mid(.Text, InStr(1, .Text, Chr$(44)) + 1, Len(.Text)))
                  If .SelStart < (Len(.Text) - hasTwoNumber) Then
                  ElseIf .SelStart >= (Len(.Text) - hasTwoNumber) Then
                     If hasTwoNumber > (FloatPoint - 1) Then KeyAscii = 0
                  End If
               Else
                  .SelText = ""
               End If
            Else
               If .SelStart > InStr(1, .Text, Chr(44)) Then KeyAscii = 0
            End If
               
            If (KeyAscii = 44) Or (KeyAscii = 46) Then
               hasPoint = InStr(1, .Text, Chr$(44))
               
               If hasPoint Then
                  KeyAscii = 0
               Else
                  If Len(.Text) < 1 Or .SelLength = Len(.Text) Then .SelText = "0"
                  KeyAscii = IIf((FloatPoint = 0), 0, 44)
               End If
            End If
         Case Else: KeyAscii = 0
      End Select
   End With
End Sub

Public Sub MaskEditMonth(ByVal ctl As TextBox, KeyAscii As Integer)
   Select Case KeyAscii
      Case 8
      Case 13: SendKey ocKEYTAB
      Case 48 To 57
         If ctl.SelStart = 2 Then ctl.SelText = "/"
      Case Else: KeyAscii = 0
   End Select
End Sub

Public Sub MaskEditDate(ByVal ctl As TextBox, KeyAscii As Integer)
   Select Case KeyAscii
      Case 8
      Case 13: SendKey ocKEYTAB
      Case 48 To 57
         If ctl.SelStart = 2 Then ctl.SelText = "/"
         If ctl.SelStart = 5 Then ctl.SelText = "/"
      Case Else: KeyAscii = 0
   End Select
End Sub

Public Sub MaskEditHour(ByVal ctl As TextBox, KeyAscii As Integer)
   Select Case KeyAscii
      Case 8
      Case 13: SendKey ocKEYTAB
      Case 48 To 57
         If ctl.SelStart = 2 Then ctl.SelText = ":"
         If ctl.SelStart = 5 Then ctl.SelText = ":"
      Case Else: KeyAscii = 0
   End Select
End Sub

'*******************************************************************
' Procedimento: MaskMoney
' Argumentos  : Ctl As TextBox
'               -> Qualquer caixa de texto para formatação
' Retorno     : Nenhum
' Objetivo    : Criar uma máscara para formatação de moeda nas caixas
'               de texto
'*******************************************************************
Public Sub MaskMoney(ByVal ctl As TextBox)
   Dim i As Integer, t As String
   
   With ctl
      t = .Text
      i = Len(t) - .SelStart
      t = Replace(.Text, ",", "")
      If Len(t) < 3 Then t = String(3 - Len(t), "0") & t
      t = Mid(t, 1, Len(t) - 2) & "," & Mid(t, Len(t) - 1)
      t = Format(t, "###,###,###,###,###,#0.00")
      If .Text <> t Then .Text = t
      .SelStart = Len(t) - i
   End With
End Sub

'*******************************************************************
' Procedimento: ContarDiasUteis
' Argumentos  : DataInicio As Date
'               -> Data de início do período da contagem
'               DataFinal As Date
'               -> Data de término do período da contagem
'               IgnorarFeriados As Boolean
'               -> Ignora os feriados existentes no período informado
' Retorno     : Long
' Objetivo    : Contar a quantidade de dias úteis num período de
'               datas informado
'*******************************************************************
Public Function ContarDiasUteis(ByVal DataInicio As Date, ByVal DataFinal As Date, ByVal IgnorarFeriados As Boolean) As Integer
   Dim i As Long
   Dim totDia As Long, iDia As Integer
   'Dim cFer As Feriado
   
   'Set cFer = New Feriado
   totDia = 0
   
   For i = DataInicio To DataFinal
      iDia = Weekday(i)
      Select Case iDia
         Case vbSunday, vbSaturday
         Case Else
            If IgnorarFeriados Then
               totDia = totDia + 1
            Else
               'If Not cFer.Existe(Format$(i, "dd/mm")) Then totDia = totDia + 1
            End If
      End Select
   Next
   
   'Set cFer = Nothing
   ContarDiasUteis = totDia
End Function

'*******************************************************************
' Procedimento: ContarSemanas
' Argumentos  : DataInicio As Date
'               -> Data de início do período da contagem
'               DataFinal As Date
'               -> Data de término do período da contagem
' Retorno     : Long
' Objetivo    : Contar a quantidade de semanas num período de
'               datas informado
'*******************************************************************
Public Function ContarSemanas(ByVal DataInicio As Date, ByVal DataFinal As Date) As Long
   Dim i As Long
   Dim totDias As Long, iDia As Integer
   Dim totSem As Long
   
   totDias = 0
   totSem = 0
   
   For i = DataInicio To DataFinal
      iDia = Weekday(i)
      Select Case iDia
         Case vbSunday, vbSaturday
            If totDias > 0 Then totSem = totSem + 1
            totDias = 0
         Case Else: totDias = totDias + 1
      End Select
   Next
   
   If totDias > 2 Then totSem = totSem + 1
   ContarSemanas = totSem
End Function

'*******************************************************************
' Procedimento: FeriadoMovel
' Argumentos  : Ano As Long
'               -> Ano de pesquisa dos feriados
' Retorno     : Date
' Objetivo    : Retorna o primeiro dia útil subsequente após/antes o
'               feriado de acordo com as opções do usuário
'*******************************************************************
Function FeriadoMovel(ByVal Ano As Long, ByRef Carnaval As Date, ByRef SextaFeiraSanta As Date, ByRef Pascoa As Date, ByRef CorpusChristi As Date) As Boolean
  On Local Error GoTo errHandle
  Dim A, B, c, d, e, f, g, h, i, k, l, M, p, q As Integer
  Dim rFeriado(1 To 4) As Date
  
  FeriadoMovel = False
  
  A = Ano Mod 19
  B = Int(Ano / 100)
  c = Ano Mod 100
  d = Int(B / 4)
  e = B Mod 4
  f = Int((B + 8) / 25)
  g = Int((B - f + 1) / 3)
  h = (19 * A + B - d - g + 15) Mod 30
  i = Int(c / 4)
  k = c Mod 4
  l = (32 + 2 * e + 2 * i - h - k) Mod 7
  M = Int((A + 11 * h + 22 * l) / 451)
  p = Int((h + l - 7 * M + 114) / 31)
  q = (h + l - 7 * M + 114) Mod 31
  
  ' *** A Páscoa será no dia Q + 1, do mês P ***
  rFeriado(1) = CDate((q + 1) & "/" & p & "/" & Ano)
  
  ' *** Carnaval: 47 dias antes da Páscoa ***
  rFeriado(2) = rFeriado(1) - 47
  
  ' *** Sexta Feira Santa (Paixão): 2 dias antes da Páscoa ***
  rFeriado(3) = rFeriado(1) - 2
  
  ' *** Corpus Christi: 60 dias após a Páscoa ***
  rFeriado(4) = rFeriado(1) + 60
  
  Carnaval = rFeriado(2)
  SextaFeiraSanta = rFeriado(3)
  Pascoa = rFeriado(1)
  CorpusChristi = rFeriado(4)
  
  FeriadoMovel = True
  Exit Function
  
errHandle:
  FeriadoMovel = False
End Function

Public Function AnoBissexto(ByVal AnoRef As Long) As Boolean
   'Um ano é bissexto se divisível por 4
   'Um ano é bissexto se divisível por 4 e não divisível por 100
   'Um ano é bissexto se divisível por 400
   
   AnoBissexto = (((AnoRef Mod 4) = 0) And ((AnoRef Mod 100) <> 0)) Or ((AnoRef Mod 400) = 0)
End Function

'*******************************************************************
' Procedimento: NumeroExtenso
' Argumentos  : Numero As Double
'               -> Número a ser escrito por extenso
'               Moeda As Double (Opcional)
'               -> Define se o número passado é valor monetário ou não
' Retorno     : String
' Objetivo    : Retorna o número escrito por extenso, incluindo
'               valores monetários
'*******************************************************************
Public Function NumeroExtenso(ByVal Numero As Double, Optional ByVal Moeda As Boolean = True) As String
   Dim i As Integer, iTam As Integer   'Declara as variáveis
   
   Dim sValor As String
   Dim sParte As String
   Dim sFinal As String
   
   'Se o número for menor que zero ou superior a 999.999.999,99 sai da função
   If Numero < 0 Or Numero > 999999999.99 Then Exit Function
   
   Dim rGrupo(4), rTexto(4) As String  'Define as variáveis
   Dim rUnidades(19) As String
   Dim rDezenas(9) As String
   Dim rCentenas(9) As String
   
   'Unidades
   rUnidades(1) = "um "
   rUnidades(2) = "dois "
   rUnidades(3) = "tres "
   rUnidades(4) = "quatro "
   rUnidades(5) = "cinco "
   rUnidades(6) = "seis "
   rUnidades(7) = "sete "
   rUnidades(8) = "oito "
   rUnidades(9) = "nove "
   rUnidades(10) = "dez "
   rUnidades(11) = "onze "
   rUnidades(12) = "doze "
   rUnidades(13) = "treze "
   rUnidades(14) = "quatorze "
   rUnidades(15) = "quinze "
   rUnidades(16) = "dezesseis "
   rUnidades(17) = "dezessete "
   rUnidades(18) = "dezoito "
   rUnidades(19) = "dezenove "
   
   'Dezenas
   rDezenas(1) = "dez "
   rDezenas(2) = "vinte "
   rDezenas(3) = "trinta "
   rDezenas(4) = "quarenta "
   rDezenas(5) = "cinquenta "
   rDezenas(6) = "sessenta "
   rDezenas(7) = "setenta "
   rDezenas(8) = "oitenta "
   rDezenas(9) = "noventa "
   
   'Centenas
   rCentenas(0) = "cem "
   rCentenas(1) = "cento "
   rCentenas(2) = "duzentos "
   rCentenas(3) = "trezentos "
   rCentenas(4) = "quatrocentos "
   rCentenas(5) = "quinhentos "
   rCentenas(6) = "seiscentos "
   rCentenas(7) = "setecentos "
   rCentenas(8) = "oitocentos "
   rCentenas(9) = "novecentos "
   
   'Formata o número para a variável local
   sValor = Format(Numero, "000000000.00")
   
   rGrupo(1) = Mid$(sValor, 1, 3)       'Divide o número em grupos de 3 dígitos
   rGrupo(2) = Mid$(sValor, 4, 3)
   rGrupo(3) = Mid$(sValor, 7, 3)
   rGrupo(4) = "0" + Mid$(sValor, 11, 2)
    
   For i = 1 To 4
      'Transfere o grupo para a variável temporária
      sParte = rGrupo(i)
      
      'Verifica o tamanho do grupo de números
      iTam = Switch(Val(sParte) < 10, 1, Val(sParte) < 100, 2, Val(sParte) < 1000, 3)
      
      'O tamanho é 3
      If iTam = 3 Then
         'Caso os 2 últimos algarismos forem 00, trata-se de uma centena inteira
         If Right(sParte, 2) <> "00" Then
            'Passa para o texto qual centena se refere
            rTexto(i) = rTexto(i) + rCentenas(Left$(sParte, 1)) + "e "
            'Diminui o tamanho para 2
            iTam = 2
         Else
            'Passa para o texto qual centena pertence o número
            rTexto(i) = rTexto(i) + IIf(Left$(sParte, 1) = "1", rCentenas(0), rCentenas(Left$(sParte, 1)))
         End If
      End If
      
      'O tamanho é 2
      If iTam = 2 Then
         'Verifica se os 2 últimos algarimos são menores que 20
         'Se positivo informa qual é a unidade a que se refere
         If Val(Right$(sParte, 2)) < 20 Then
            'Adiciona ao texto a unidade referente ao número
            rTexto(i) = rTexto(i) + rUnidades(Right$(sParte, 2))
         Else
            'Adicona ao texto a dezena referente ao número
            rTexto(i) = rTexto(i) + rDezenas(Mid$(sParte, 2, 1))
            
            'Caso não seja uma dezena exata
            If Right$(sParte, 1) <> "0" Then
               'Adicona ao texto a palavra 'e'
               rTexto(i) = rTexto(i) + "e "
               'Diminui o tamanho para 1
               iTam = 1
            End If
         End If
      End If
      
      'O tamanho é 1
      If iTam = 1 Then
         'Adicona ao texto qual unidade represena o número
         rTexto(i) = rTexto(i) + rUnidades(Right$(sParte, 1))
      End If
   Next
   
   'Se é moeda, verifica se possui centavos
   If Val(rGrupo(1) + rGrupo(2) + rGrupo(3)) = 0 And Val(rGrupo(4)) <> 0 Then
      'Se é 1, então adiciona o texto no singular, senão, no plural
      sFinal = rTexto(4) + IIf(Val(rGrupo(4)) = 1, "centavo", "centavos")
   Else
      'Limpa a variável
      sFinal = ""
      
      'Verifica se o grupo dos milhões possui unidades, se positivo
      'adicona a palava milhão ou milhões conforme a unidade for 1 ou maior
      sFinal = sFinal + IIf(Val(rGrupo(1)) <> 0, rTexto(1) + IIf(Val(rGrupo(1)) > 1, "milhões ", "milhão "), "")
      
      'Verifica se os grupos 2 e 3 são 0
      If Val(rGrupo(2) + rGrupo(3)) = 0 Then
         'Se positivo, acrescenta a preposição 'de'
         sFinal = sFinal + "de "
      Else
         'Se negativo, verifica se o grupo 2 é maior que 0 e
         'adiciona a palavar mil
         sFinal = sFinal + IIf(Val(rGrupo(2)) <> 0, rTexto(2) + "mil " & IIf(Val(rGrupo(3)) < 100, "e ", ""), "")
      End If
      
      'Se o número não se trata de valor monetário
      If Not Moeda Then
         'Adicona a variável o grupo 3 e o grupo 4, caso este seja maior que 0
         sFinal = sFinal + rTexto(3) + IIf(Val(rGrupo(4)) <> 0, ", " + rTexto(4), "")
      Else
         'Adiciona a variável o grupo 3 e verifica se os grupos 1, 2 e 3 são maiores que 1
         'Se positivo, adiciona a palavra reais; se negativo, a palavra real
         sFinal = sFinal + rTexto(3) + IIf(Val(rGrupo(1) + rGrupo(2) + rGrupo(3)) = 1, "real ", "reais ")
         
         'Adiciona a variavel o grupo 4 se maior que 0, e verifica se é maior que 1
         'Se positivo, adicona a palavra centavo; se negativo, a palavra centavos
         'Se igual a 0, não adiciona nada
         sFinal = sFinal + IIf(Val(rGrupo(4)) <> 0, "e " + rTexto(4) + IIf(Val(rGrupo(4)) = 1, "centavo", "centavos"), "")
        End If
   End If
   
   'Retorna o número por extenso
   NumeroExtenso = sFinal
End Function

Public Function ValidateNull(ByVal var As ADODB.Field) As Variant
   Select Case var.Type
      Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adTinyInt
         ValidateNull = IIf(IsNull(var), 0, var)
      Case adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
         ValidateNull = IIf(IsNull(var), 0, var)
      Case adDate, adDBDate, adDBTime, adDBTimeStamp
         ValidateNull = IIf(IsNull(var), 0, var)
      Case adBoolean
         ValidateNull = IIf(IsNull(var), False, var)
      Case adBSTR, adChar, adLongVarBinary, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
         ValidateNull = IIf(IsNull(var), "", var)
      Case Else
         ValidateNull = IIf(IsNull(var), "", var)
   End Select
End Function

Public Function SetValidateNull(ByVal var As Variant) As Variant
   Select Case VarType(var)
      Case vbByte
         SetValidateNull = IIf((CByte(var) = 0), "Null", CByte(var))
      Case vbInteger
         SetValidateNull = IIf((CInt(var) = 0), "Null", CInt(var))
      Case vbLong
         SetValidateNull = IIf((CLng(var) = 0), "Null", CLng(var))
      Case vbCurrency
         SetValidateNull = IIf((CCur(var) = 0), "Null", CCur(var))
      Case vbDecimal
         SetValidateNull = IIf((CDec(var) = 0), "Null", CDec(var))
      Case vbDouble
         SetValidateNull = IIf((CDbl(var) = 0), "Null", CDbl(var))
      Case vbSingle
         SetValidateNull = IIf((CSng(var) = 0), "Null", CSng(var))
      Case vbDate
         SetValidateNull = IIf((CDate(var) = 0), "Null", CDate(var))
      Case vbNull
         SetValidateNull = "Null"
      Case vbString
         SetValidateNull = IIf((CStr(var) = 0), "Null", CStr(var))
      Case vbVariant
         SetValidateNull = IIf((CVar(var) = 0), "Null", CVar(var))
      Case Else
         SetValidateNull = "Null"
   End Select
End Function

'*******************************************************************
' Procedimento: RemoverAcento
' Argumentos  : Texto As String
'               -> Texto a ser validado
' Retorno     : String
' Objetivo    : Remover os acentos do texto especificado
'*******************************************************************
Public Function RemoverAcento(ByVal Texto As String) As String
   Dim i As Integer, p As Integer, aux As String
   Dim Txt As String
   
   Const t1 = "ÀàÈèÌìÒòÙùÁáÉéÍíÓóÚúÂâÊêÎîÔôÛûÃãÕõÄäËëÏïÖöÜüÑñÇçªº°"
   Const t2 = "AaEeIiOoUuAaEeIiOoUuAaEeIiOoUuAaOoAaEeIiOoUuNnCcaoo"
   
   p = 1
   Txt = Texto
   
   For i = 1 To Len(Txt)
      aux = Mid$(Txt, i, 1)    'pega o caracter
      p = InStr(1, t1, aux)    'tem acento correpondente?
      If p Then                'tem...
         'Mid(Txt, i, 1) = aux    'Método rápido
         Txt = Mid$(Txt, 1, i - 1) & Mid$(t2, p, 1) & Mid$(Txt, i + 1) 'troca pelo correpondente
      End If
   Next
   
   RemoverAcento = Txt
End Function

Public Function RemoverFormato(ByVal TextoFormatado As String) As String
   Dim i As Integer
   Dim NewString As String
   Dim AscChar As Integer
   
   NewString = vbNullString
   
   For i = 1 To Len(TextoFormatado)
      AscChar = Asc(Mid$(TextoFormatado, i, 1))
      Select Case AscChar
         Case 40, 41, 44 To 47, 58, 64
         Case 48 To 57, 65 To 90, 97 To 122
            NewString = NewString & Mid$(TextoFormatado, i, 1)
      End Select
   Next
   
   RemoverFormato = NewString
End Function

Public Function ResolverString(ByVal resString As String, ParamArray varReplacements() As Variant) As String
   Dim intMacro As Integer
   Dim strResString As String
   
   Dim strMacro As String
   Dim strValue As String
   Dim intPos As Integer
   
   strResString = resString
   
   For intMacro = LBound(varReplacements) To UBound(varReplacements) Step 2
      strMacro = varReplacements(intMacro)
      On Error GoTo MismatchedPairs
      strValue = varReplacements(intMacro + 1)
      On Error GoTo 0
        
      Do
         intPos = InStr(1, strResString, strMacro)
         If intPos > 0 Then
            strResString = Left$(strResString, intPos - 1) & strValue & Right$(strResString, Len(strResString) - Len(strMacro) - intPos + 1)
         End If
      Loop Until intPos = 0
   Next intMacro
   
   ResolverString = strResString
   Exit Function
    
MismatchedPairs:
   Resume Next
End Function

'*******************************************************************
' Procedimento: ValidarEMail
' Argumentos  : EnderecoEMail As String
'               -> endereço do email
' Retorno     : Boolean
' Objetivo    : Valida se o endereço de e-mail é válido
'*******************************************************************
Public Function ValidarEMail(ByVal EnderecoEMail As String) As Boolean
   Dim nCharacter As Integer
   Dim Count As Integer
   Dim sLetra As String
   
   'Atribui falha na execução
   ValidarEMail = False
   
   'Verifica se o e-mail tem no MÍNIMO 5 caracteres (a@b.c)
   If Len(EnderecoEMail) < 5 Then
      'O e-mail é inválido, pois tem menos de 5 caracteres
      ShowMsg "O e-mail informado tem menos de 5 caracteres !!!", vbCritical
      Exit Function
   End If
   
   'Verificar a existencia de arrobas (@) no e-mail
   For nCharacter = 1 To Len(EnderecoEMail)
      If Mid$(EnderecoEMail, nCharacter, 1) = "@" Then
         'OPA!!! Achou uma arroba!!!
         'Soma 1 ao contador
         Count = Count + 1
      End If
   Next
   
   'Verifica o número de arrobas.
   'TEM que ter """UMA""" arroba
   If Count <> 1 Then 'O e-mail é inválido, pois tem 0 ou mais de 1 arroba
      ShowMsg "O número de arrobas (@) do e-mail é inválido !!!", vbCritical
      Exit Function
   End If
   
   'O e-mail tem 1 arroba.
   'Verificar a posição da arroba
   If InStr(EnderecoEMail, "@") = 1 Then
      'O e-mail é inválido, pois começa com uma @
      ShowMsg "O e-mail foi iniciado com uma arroba (@) !!!", vbCritical
      Exit Function
   ElseIf InStr(EnderecoEMail, "@") = Len(EnderecoEMail) Then
      'O e-mail é inválido, pois termina com uma @
      ShowMsg "O e-mail termina com uma arroba (@) !!!", vbCritical
      Exit Function
   End If
   
   nCharacter = 0
   Count = 0
   
   'Verificar a existencia de pontos (.) no e-mail
   For nCharacter = 1 To Len(EnderecoEMail)
      If Mid(EnderecoEMail, nCharacter, 1) = "." Then
         'Soma 1 ao contador
         Count = Count + 1
      End If
   Next
   
   'Verifica o número de pontos.
   'TEM que ter PELO MENOS UM ponto.
   If Count < 1 Then
      'O e-mail é inválido, pois não tem pontos.
      ShowMsg "O e-mail é inválido, pois não contém pontos (.) !!!", vbCritical
      Exit Function
   End If
   
   'O e-mail tem pelo menos 1 ponto.
   'Verificar a posição do ponto:
   If InStr(EnderecoEMail, ".") = 1 Then
      'O e-mail é inválido, pois começa com um ponto
      ShowMsg "O e-mail foi iniciado com um ponto (.) !!!", vbCritical
      Exit Function
   ElseIf InStr(EnderecoEMail, ".") = Len(EnderecoEMail) Then
      'O e-mail é inválido, pois termina com um ponto.
      ShowMsg "O e-mail termina com um ponto (.) !!!", vbCritical
      Exit Function
   ElseIf InStr(InStr(EnderecoEMail, "@"), EnderecoEMail, ".") = 0 Then
      'O e-mail é inválido, pois não possui ponto após a arroba
      ShowMsg "O e-mail não tem nenhum ponto (.) após a arroba (@) !!!", vbCritical
      Exit Function
   End If
   
   nCharacter = 0
   Count = 0
   
   'Verifica se o e-mail não tem pontos
   'consecutivos (..) após a arroba (@).
   If InStr(EnderecoEMail, "..") > InStr(EnderecoEMail, "@") Then
      'O e-mail é inválido, tem pontos consecutivos após o @.
      ShowMsg "O e-mail contém pontos consecutivos (..) após o arroba (@) !!!", vbCritical
      Exit Function
   End If
   
   'Verifica se o e-mail tem caracteres
   'inválidos
   For nCharacter = 1 To Len(EnderecoEMail)
      sLetra = Mid$(EnderecoEMail, nCharacter, 1)
      If Not (LCase(sLetra) Like "[a-z]" Or sLetra = "@" Or sLetra = "." Or sLetra = "-" Or sLetra = "_" Or IsNumeric(sLetra)) Then
         'O e-mail é inválido, pois tem caracteres inválidos
         ShowMsg "Foi digitado um caracter inválido no e-mail !!!", vbCritical
         Exit Function
      End If
   Next
   
   nCharacter = 0
   
   'Bem, se a verificação chegou até aqui
   'é porque o e-mail é válido, então...
   ValidarEMail = True
End Function

'*******************************************************************
' Procedimento: ValidarCMC7
' Argumentos  : Numero As String
'               -> Código de barras do cheque
' Retorno     : Boolean
' Objetivo    : Valida se o código de barras do cheque especificado
'               é válido
'*******************************************************************
Public Function ValidarCMC7(ByVal Numero As String) As Boolean
   Dim b1 As Boolean, b2 As Boolean, b3 As Boolean
   Dim c1 As String, c2 As String, c3 As String
   Dim d1 As String, d2 As String, d3 As String
   
   c1 = Mid$(Numero, 1, 7)
   c2 = Mid$(Numero, 9, 10)
   c3 = Mid$(Numero, 20, 10)
   
   d1 = Mid$(Numero, 19, 1)
   d2 = Mid$(Numero, 8, 1)
   d3 = Mid$(Numero, 30, 1)

   b1 = (DVBase10(c1) = d1)
   b2 = (DVBase10(c2) = d2)
   b3 = (DVBase10(c3) = d3)
   
   ValidarCMC7 = b1 And b2 And b3
End Function

'*******************************************************************
' Procedimento: DVBase10
' Argumentos  : Codigo As String
'               -> Código a ser verificado
' Retorno     : Integer
' Objetivo    : Retorna o dígito verificado da base 10
'*******************************************************************
Private Function DVBase10(ByVal Codigo As String) As Integer
   Dim bFlag As Boolean
   Dim i As Integer
   Dim DV As Long
   Dim dig As Integer
   
   DV = 0
   bFlag = True
  
   For i = Len(Codigo) To 1 Step -1
      If bFlag Then
         dig = CInt(Mid$(Codigo, i, 1)) * 2
      Else
         dig = CInt(Mid$(Codigo, i, 1))
      End If
      
      bFlag = Not bFlag
      
      If dig > 9 Then
         dig = 1 + (dig - 10)
         DV = DV + dig
      Else
         DV = DV + dig
      End If
   Next
   
   dig = 10 * ((DV / 10) - Int(DV / 10))
   If dig > 0 Then dig = 10 - dig
   
   DVBase10 = dig
End Function

'*******************************************************************
' Procedimento: GetDefalutPrinter
' Argumentos  : Nenhum
' Retorno     : Object
'               -> Retorna um objeto Printer
' Objetivo    : Recupera o nome da impressora padrão do sistema
'*******************************************************************
Public Function GetDefaultPrinter() As Object
   'Declara as variáveis
   Dim strBuffer As String * 254
   Dim iRetValue As Long
   Dim strDefaultPrinterInfo As String
   Dim tblDefaultPrinterInfo() As String
   Dim objPrinter As Printer
   
   'Executa a função que retorna a impressora padrão do Windows
   iRetValue = GetProfileString("windows", "device", ",,,", strBuffer, 254)
   
   'Realiza tratamento na variável de retorno
   strDefaultPrinterInfo = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
   
   'Cria uma matriz com as informações da impressora
   tblDefaultPrinterInfo = Split(strDefaultPrinterInfo, ",")
   
   'Percorre todas as impressoras instaladas
   For Each objPrinter In Printers
      'Se encontrou a impressora, sai do laço
      If objPrinter.DeviceName = tblDefaultPrinterInfo(0) Then Exit For
   Next
   
   'Se impressora diferente do padrão destrói a variável
   If objPrinter.DeviceName <> tblDefaultPrinterInfo(0) Then Set objPrinter = Nothing
   
   'Retorna o resultado
   Set GetDefaultPrinter = objPrinter
End Function

Public Function GerarCodigoDisponivel(ByVal Tabela As String, ByVal CampoPesquisa As String, Optional ByVal Grupo As Long = 0, Optional ByVal CodigoIni As Long = 0) As String
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim bRet As Boolean
   Dim codLivre As Long
   Dim sRet As String
   
   'Atribui o código inicial
   codLivre = CodigoIni
   
   Do
      codLivre = codLivre + 1
      sSQL = "SELECT " & CampoPesquisa & " FROM " & Tabela & " WHERE (" & CampoPesquisa & " = " & codLivre & ");"
      Set r = dbData.OpenRecordset(sSQL)
      bRet = Not r.BOF
      If r.State <> 0 Then r.Close
      Set r = Nothing
   Loop While bRet
   
   sRet = ""
   If Grupo = 0 Then
      sRet = codLivre
   Else
      sRet = Format$(Grupo, "00")
      sRet = sRet & Right$(codLivre, 4)
   End If
   
   GerarCodigoDisponivel = sRet
End Function

Public Function GetFileName(ByVal FullPath As String, Optional ByVal Extension As Boolean = True) As String
   Dim iPos As Integer
   Dim fName As String
   
   iPos = InStrRev(FullPath, "\")
   fName = Mid$(FullPath, iPos + 1)
   
   If (Not Extension) Then
      fName = Mid$(fName, 1, InStrRev(fName, ".") - 1)
   End If
   
   GetFileName = fName
End Function

'*******************************************************************
' Procedimento: NormalizePath
' Argumentos  : FullPath As String
'               -> Caminho do diretório
' Retorno     : FullPath As String
' Objetivo    : Retorna o caminho do diretório acrescido de \
'*******************************************************************
Public Sub NormalizePath(ByRef FullPath As String)
   If Right$(FullPath, 1) <> "\" Then FullPath = FullPath & "\"
End Sub

Public Sub VerifyPathTree(ByVal FullPath As String)
   Dim sPath As String, newPath As String
   Dim i As Integer
   
   newPath = ""            'Inicializa o novo diretório
   sPath = FullPath        'Atribui o diretório especificado
   NormalizePath sPath     'Normaliza o diretório
   
   If Dir$(sPath, vbDirectory) = "" Then
      While InStr(2, sPath, "\") > 0
         newPath = newPath & Left(sPath, InStr(2, sPath, "\") - 1)
         sPath = Mid$(sPath, InStr(2, sPath, "\"))
         If Dir$(newPath, 16) = "" Then
            Debug.Print newPath
            MkDir newPath
         End If
      Wend
   End If
End Sub

'*******************************************************************
' Procedimento: VerifyDateTree
' Argumentos  : CheckDate As Date
'               -> Data para checagem do diretório
'               DirBase As String
'               -> Diretório base da árvore
' Retorno     : Nenhum
' Objetivo    : Verifica a árvore de diretórios e cria as pastas
'               se não exisitirem. Utiliza o padrão AAAA\MM\DD
'*******************************************************************
Public Sub VerifyDateTree(ByVal CheckDate As Date, ByVal DirBase As String)
   Dim rDir() As String
   Dim sDrv As String, sPath As String, newPath As String
   Dim i As Integer
   
   sDrv = DirBase
   NormalizePath sDrv
   
   sPath = Format$(CheckDate, "yyyy\\mm")  'Format$(CheckDate, "yyyy") & "\" & Format$(CheckDate, "mm") & "\" & Format$(CheckDate, "dd")
   NormalizePath sPath
   
   newPath = ""
   If Dir$(sDrv & sPath, vbDirectory) = "" Then
      While InStr(2, sPath, "\") > 0
         newPath = newPath & Left(sPath, InStr(2, sPath, "\") - 1)
         sPath = Mid$(sPath, InStr(2, sPath, "\"))
         If Dir$(sDrv & newPath, 16) = "" Then
            Debug.Print sDrv & newPath
            MkDir sDrv & newPath
         End If
      Wend
   End If
End Sub

'*******************************************************************
' Procedimento: DateTree
' Argumentos  : DirDate As Date
'               -> Data para transformação em diretórios
' Retorno     : String
' Objetivo    : Retorna o formato do diretório pela data informada
'               Utiliza o formato AAAA\MM\DD
'*******************************************************************
Public Function DateTree(ByVal DirDate As Date) As String
   DateTree = Format$(DirDate, "yyyy") & "." & Format$(DirDate, "mm") & "." & Format$(DirDate, "dd")
End Function

Public Sub SelectControl(ByVal rControl As Object)
   Const SEL_FLD = "O controle não possui uma propriedade texto para que possa ser selecionado."
   
   If TypeOf rControl Is TextBox Then GoTo SelectField
   If TypeOf rControl Is ComboBox Then
      If rControl.Style <> 2 Then GoTo SelectField
   End If
   If TypeOf rControl Is MaskEdBox Then GoTo SelectField
   'If TypeOf rControl Is PickBox Then GoTo SelectField
   
   MsgBox SEL_FLD, vbExclamation
   Exit Sub

SelectField:
   If Len(rControl.Text) = 0 Then Exit Sub
   rControl.SelStart = 0
   rControl.SelLength = Len(rControl.Text)
End Sub

'*******************************************************************
' Procedimento: SelecionarValorNaLista
' Argumentos  : Lista As Object
'               -> O objeto pode ser uma ListBox ou ComboBox
'               Valor As Long
'               -> Valor a ser selecionado dentro da lista
' Retorno     : Nenhum
' Objetivo    : Percorre cada item da lista, verificando o conteúdo
'               da propriedade ItemData. Caso o valo da propriedade
'               seja igual a variável Valor, seleciona o item
'               correspondente
'*******************************************************************
Public Sub SelecionarValorNaLista(Lista As Object, ByVal Valor As Long)
   'Inicia o controle de erro
   On Local Error GoTo errHandle
   Dim i As Integer        'Declara as variáveis
   
   'Remove qualquer seleção da lista
   Lista.ListIndex = -1
   
   'Executa a verificação do primeiro ao último item da lista
   For i = 0 To Lista.ListCount - 1
      'Se ItemData = Valor
      If Lista.ItemData(i) = Valor Then
         Lista.ListIndex = i  'Seleciona o item
         Exit For             'Sai do loop
      End If
   Next
   Exit Sub       'Sai da rotina
   
errHandle:
   'Em caso de erro, não seleciona nenhum item
   Lista.ListIndex = -1
End Sub

Public Function MontarCriterios(ByVal Criterio As String) As String
   On Local Error GoTo errHandle
   Dim i As Integer, j As Integer
   Dim novoCriterio As String
   Dim rVal1() As String, rVal2() As String
   
   novoCriterio = ""
   rVal1() = Split(Criterio, ";")
   
   For i = LBound(rVal1) To UBound(rVal1)
      If InStr(1, rVal1(i), "-") > 0 Then
         rVal2() = Split(rVal1(i), "-")
         
         For j = CLng(rVal2(0)) To CLng(rVal2(1))
            novoCriterio = novoCriterio & "'" & j & "', "
         Next
      Else
         novoCriterio = novoCriterio & "'" & rVal1(i) & "', "
      End If
   Next
   
   novoCriterio = Left$(Trim(novoCriterio), Len(Trim(novoCriterio)) - 1)
   MontarCriterios = novoCriterio
   Exit Function
   
errHandle:
   If Err.Number = 13 Then
      MontarCriterios = "#1"
   Else
      MontarCriterios = ""
   End If
End Function

Public Function ExistInList(ByVal objList As Object) As Boolean
   Dim i As Integer, bExiste As Boolean
   
   If objList.ListCount = 0 Then
      ExistInList = True
      Exit Function
   End If
   
   bExiste = False
   For i = 0 To objList.ListCount - 1
      If LCase(objList) = LCase(objList.List(i)) Then
         bExiste = True
         Exit For
      End If
   Next
   
   ExistInList = bExiste
End Function

Public Function CriarChaveLancamento(ByVal randomNumber As Long) As String
   On Local Error GoTo errHandle
   Dim i As Integer, n As Integer
   Dim sKey As String
   
   Const ALFABETO = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
   Const LEN_ALFA = 36
   
   'Cria a nova ordem aleatória
   Randomize randomNumber
   
   'Inicializa a chave
   sKey = ""
   
   'Processa as iterações necessárias
   For i = 1 To 12
      'Seleciona um numero randômico
      n = Int((LEN_ALFA * Rnd) + 1)
      
      'Atribui a letra escolhida do alfabeto
      sKey = sKey & Mid$(ALFABETO, n, 1)
   Next
   
   'Retorna o resultado
   CriarChaveLancamento = sKey
   Exit Function
   
errHandle:
   CriarChaveLancamento = ""
End Function

'abre banco de dados
Public Function AbreBancoDeDados(Optional ByVal vgIgnoraErros As Long = 0) As Integer
    Dim x As String
    On Error GoTo deuErro
    AbreBancoDeDados = False
    x$ = "Provider=SQLOLEDB.1;Persist Security Info=False;DRIVER={Sql Server};SERVER=" + vgServerName + ";uid=sa;pwd=190106web;DATABASE=cyber_base;TRUSTED_CONNECTION=NO"
    'x$ = "Provider=SQLOLEDB.1;Persist Security Info=False;DRIVER={Sql Server};SERVER=" + vgServerName + ";uid=lotesis;pwd=lotesis;DATABASE=cyber_base;TRUSTED_CONNECTION=NO"
    vgDb.ConnectionString = x$
    vgDb.CursorLocation = adUseClient
    vgDb.Open
    vgCat.ActiveConnection = vgDb

    AbreBancoDeDados = True                       'tudo certo!
    Exit Function

deuErro:
 If vgIgnoraErros = 0 Then
   Screen.MousePointer = vbDefault                     'cursor = seta
   'Prepara String, para buscar os dados
   x$ = Replace(x$, ";", "];[")
  'Se o erro eh de conexão com o BD mostra para o usuário a mensagem mais amigável
  If Err = -2147467259 Then
    MsgBox "Falha de conexão com o servidor SQL-Server!" & vbCrLf & _
           "  Servidor: " & DadosExtrasSepara(x$, "SERVER", "") & vbCrLf & _
           "  Banco...: " & DadosExtrasSepara(x$, "DATABASE", "") & vbCrLf & vbCrLf & _
           "Importante: Verifique o servidor está ligado ou se o serviço do SQL-Server está ativo!" & vbCrLf & _
           "Em caso de dúvida contacte o suporte técnico!", vbCritical
    End
  End If
  'Esse é o padrao
  MsgBox Err.Number + " - " + Err.Description, vbCritical
  Err.Clear
 End If
End Function

'fecha o banco de dados
Public Sub FechaBancoDeDados()
    On Error Resume Next

'    'fecha arquivo de controle de cps sequenciais
'    If Not vgRsSequencia Is Nothing Then
'        vgRsSequencia.Close
'        Set vgRsSequencia = Nothing
'    End If
'
'    'fecha arquivos de perametros
'    If Not TbContabilista Is Nothing Then
'        TbContabilista.Close
'        Set TbContabilista = Nothing
'    End If
'    If Not TbEmpresa Is Nothing Then
'        TbEmpresa.Close
'        Set TbEmpresa = Nothing
'    End If

    If Not vgDb Is Nothing Then
        Set vgCat = Nothing
        vgDb.Close                                'fecha o banco de dados
        Set vgDb = Nothing                        'libera memória
    End If
End Sub

'abre tabela no SQL
Public Sub RsOpen(ByRef vgRs As ADODB.Recordset, ByVal vgSQL As String)
    Dim i As Long, j As Long
    Dim x As String, z As String, zz As String, xx As String

    On Error Resume Next                          'previne erro
    vgRs.Close                                    'tenta fechar a tabela
    If Err Then                                   'se não consegui
        Err.Clear                                 'tira o erro
    End If
    If Not TypeOf vgRs Is ADODB.Recordset Then    'se não for um recordset
        Set vgRs = New ADODB.Recordset            'válido tenta criar um
    End If
    On Error GoTo 0                               'se der erro passar para quem chamou

    'vamos corrigir prováveis querys incorretas que referenciam os campos com ! ao invéz de .
    j = 1
    Do
        i = InStr(j, vgSQL, "!")
        If i Then
            If (Tally(Left(vgSQL, i - 1), Chr(34)) And 1) = 0 And (Tally(Left(vgSQL, i - 1), Chr(39)) And 1) = 0 Then
                Mid(vgSQL, i, 1) = "."
            End If
            j = i + 1
        End If
    Loop Until i = 0

    'Vamos fazer algumas correões na query para SQL ou ORACLE
    vgSQL$ = Substitui(vgSQL$, "= NULL", "IS NULL", SO_UM)
    vgSQL$ = Substitui(vgSQL$, "= 'NULL'", "IS NULL", SO_UM)

    z$ = "=FALSE |=FALSE)|= FALSE |= FALSE)|=TRUE |=TRUE)|= TRUE |= TRUE)"
    x$ = ExtraiSQL(vgSQL$, EXP_WHERE)
    Do While Len(z$)
        xx$ = Parse$(z$, "|")
        zz$ = ""
        If Right(xx$, 1) = ")" Then zz$ = ")"
        If InStr(UCase(xx$), "TRUE") Then
            x$ = Substitui$(x$, xx$, " = 1 " + zz$, SO_UM)
        Else
            x$ = Substitui$(x$, xx$, " = 0 " + zz$, SO_UM)
        End If
    Loop
    vgSQL$ = InsereSQL(vgSQL$, EXP_WHERE, x$)

    'abre o recordset
AbreRecordset:
    vgRs.Open vgSQL$, vgDb, adOpenDynamic, adLockPessimistic, adCmdText
    vgRs.Properties("Update Criteria").Value = adCriteriaKey
End Sub

'Funcao que separa os dados do correntista que está no memo
Public Function DadosExtrasSepara(Memo As String, Campo As String, Optional Padrao As Variant) As String
  Dim pi As Integer, pf As String
  On Error GoTo deuErro
  pi = InStr(1, UCase(Memo), "[" & UCase(Campo) & "=")
  pf = InStr(pi, UCase(Memo), "]")
  pi = pi + Len(Campo) + 2
  If (pf - pi > 0) Then
    DadosExtrasSepara = Mid(Memo, pi, pf - pi)
  Else
    DadosExtrasSepara = Padrao
  End If
  Exit Function
  
deuErro:
  Err.Clear
  DadosExtrasSepara = Padrao
End Function

'conta quanto vezes uma string aparece em uma outra
Public Function Tally(vgAlvo As String, vgOq As String) As Integer
    Dim i As Long, vgQt As Integer
    vgQt = 0                                      'inicializa variaveis
    i = 0

OutraVez:
    i = InStr(i + 1, vgAlvo$, vgOq$)              'procura...
    If i > 0 Then                                 'se achou
        vgQt = vgQt + 1                           'soma a quantidade
        GoTo OutraVez                             'e procura mais
    End If
    Tally = vgQt
End Function

'insere uma nova cláusula na expressão SQL
Public Function InsereSQL(ByVal vgExpSQL As String, ByVal vgQual As Integer, ByVal vgOQueInserir As String) As String
    Dim vgRetVal As String, i As Integer, x As String, vgExpTop As String
    vgRetVal$ = ""                                     'conter toda a exp SQL
    vgOQueInserir$ = Trim$(vgOQueInserir$)             'cláusula a inserir
    vgExpTop$ = ExtraiSQL(vgExpSQL$, EXP_SELECT)
    If UCase$(Left$(vgExpTop$, 4)) = "TOP " Or UCase$(Left$(vgExpTop$, 8)) = "PERCENT " Then
        x$ = Parse(vgExpTop$, Chr(32))
        x$ = x$ + Chr(32) + Parse(vgExpTop$, Chr(32))
        If Left(vgExpTop$, 7) = "PERCENT" Then
            x$ = x$ + Chr(32) + "PERCENT"
        End If
        vgExpTop$ = x$
    Else
        vgExpTop$ = ""
    End If
    For i = 0 To EXP_TODAS - 1                         'corre todas as cláusulas
        If i = vgQual Then                             'se for a que quer inserir
            x$ = vgOQueInserir$                        'substitui pela informada
        Else                                           'caso contrário
            x$ = ExtraiSQL$(vgExpSQL$, i, True)        'tira cláusula da própria exp SQL
        End If
        If Len(x$) Then                                'se a clusula existe segue montando nova exp SQL
            vgRetVal$ = vgRetVal$ + LTrim$(vgClausula$(i)) + x$ + vbCrLf
        End If
    Next
    If Len(vgExpTop$) And vgQual <> EXP_SELECT Then
        x$ = ExtraiSQL(vgRetVal$, EXP_SELECT)
        If UCase(Left(x$, 3)) <> "TOP" And UCase(Left(x$, 7)) <> "PERCENT" Then
            x$ = vgExpTop$ + Chr(32) + x$
            vgRetVal$ = InsereSQL(vgRetVal$, EXP_SELECT, x$)
        End If
    End If
    InsereSQL = Trim$(vgRetVal$)                       'esta é a nova exp SQL
End Function

'Extrai a clausula escolhida da expressao SQL
Public Function ExtraiSQL(ByVal vgExpSQL As String, ByVal vgQualSQL As Integer, Optional vgTiraTop As Variant) As String
    Dim vgPosIni As Integer, vgPosFim As Integer, x As String
    Dim vgExpNormal As String, vgExpMaiusc As String, i As Integer, j As Integer
    vgExpNormal$ = " " + vgExpSQL$ + " "
    vgExpNormal$ = Substitui$(vgExpNormal$, Chr$(13), " ", SO_UM)
    vgExpNormal$ = Substitui$(vgExpNormal$, Chr$(10), " ", SO_UM)
    vgExpNormal$ = Substitui$(vgExpNormal$, " ,", ",", SO_UM)
    vgExpNormal$ = Substitui$(vgExpNormal$, "  ", " ", SO_UM)
    vgExpMaiusc$ = UCase$(vgExpNormal$)
    vgPosIni = InStr(vgExpMaiusc$, vgClausula$(vgQualSQL))
    If vgPosIni > 0 Then
        Do While vgPosIni > 0 And (Tally(Left$(vgExpMaiusc$, vgPosIni), "(") <> Tally(Left$(vgExpMaiusc$, vgPosIni), ")") Or _
           Tally(Left$(vgExpMaiusc$, vgPosIni), "[") <> Tally(Left$(vgExpMaiusc$, vgPosIni), "]"))
            vgPosIni = InStr(vgPosIni + 4, vgExpMaiusc$, vgClausula$(vgQualSQL))
        Loop
    End If
    If vgPosIni > 0 Then
        vgPosIni = vgPosIni + Len(vgClausula$(vgQualSQL))
        vgPosFim = Len(vgExpMaiusc$)
        For i = 0 To EXP_TODAS - 1
            j = InStr(vgExpMaiusc$, vgClausula$(i))
            Do While j > 0 And (Tally(Left$(vgExpMaiusc$, j), "(") <> Tally(Left$(vgExpMaiusc$, j), ")") Or _
               Tally(Left$(vgExpMaiusc$, j), "[") <> Tally(Left$(vgExpMaiusc$, j), "]"))
                j = InStr(j + 4, vgExpMaiusc$, vgClausula$(i))
            Loop
            If j > vgPosIni And j < vgPosFim Then vgPosFim = j
        Next
        x$ = Trim$(Mid$(vgExpNormal$, vgPosIni, (vgPosFim - vgPosIni) + 1))
        If Not IsMissing(vgTiraTop) Then
            If vgQualSQL = EXP_SELECT And vgTiraTop Then  'extrai o TOP n PERCENT
                If UCase$(Left$(x$, 4)) = "TOP " Then
                    x$ = LTrim$(Mid$(x$, 5))
                    If Val(x$) > 0 Then x$ = LTrim$(Mid$(x$, InStr(x$, " ")))
                    If UCase$(Left$(x$, 8)) = "PERCENT " Then
                        x$ = LTrim$(Mid$(x$, 9))
                    End If
                End If
            End If
        End If
    Else
        If vgQualSQL = EXP_FROM And Len(vgExpSQL$) > 0 Then 'so tem tabela
            x$ = vgExpSQL$
        Else
            x$ = ""
        End If
    End If
    ExtraiSQL = x$
End Function


'remove caracteres de uma string
Public Function Retira(vgAlvo As String, vgOQue As String, Como As Integer) As String
    Dim x As String, k As String, i As Integer, _
    p As Integer                                         'dimensiona
    If Como = UM_A_UM Then                               'se um a um
        x$ = ""                                          'vamos concatenar em x
        For i = 1 To Len(vgAlvo$)                        'cada caracter que
            k$ = Mid$(vgAlvo$, i, 1)                     'não estiver
            If InStr(vgOQue$, k$) = 0 Then x$ = x$ + k$  'contido na string a regirar
        Next
    Else                                                 'se não for um a um
        x$ = vgAlvo$                                     'vamos tirar

ProcuraOutro:
        p = InStr(x$, vgOQue$)                           'toda a string
        If p > 0 Then                                    'de uma só vez
            x$ = Left$(x$, p - 1) + Mid$(x$, p + Len(vgOQue$)) 'da string alvo
            GoTo ProcuraOutro
        End If
    End If
    Retira$ = x$                                               'retorna nova string
End Function

'troca caracter por outro, dentro da string
Public Function Substitui(vgAlvo As String, vgOQue As String, vgPeloQue As String, Como As Integer) As String
    Dim x As String, k As String, p As Long, i As Integer       'dimensiona
    x$ = vgAlvo$                                                'salva string alvo
    If Como = UM_A_UM Then                                      'se um a um,
        p = 1
        For i = 1 To Len(vgOQue$)                               'vamos trocar
            k$ = Mid$(vgOQue$, i, 1)                            'cada caracter de vgOQue$
            p = InStr(p, x$, k$)                                'pelo correspondente em vgPeloQue$
            If p > 0 Then                                       'caracter encontrado
                Mid$(x$, p, 1) = Mid$(vgPeloQue$, i, 1)         'substitui na string alvo
                p = p + 1                                       'vamos contiuar procurando
                i = i - 1                                       'o mesmo caracter
            Else
                p = 1                                           'prepara para pesquisar o proximo caracter
            End If
        Next
    Else                                          'senão,
        p = InStr(UCase(x$), UCase(vgOQue$))      'vamos trocar
        While p > 0                               'todos de uma vez
            x$ = Left$(x$, p - 1) + vgPeloQue$ + Mid$(x$, p + Len(vgOQue$)) 'quantas vezes necessário
            p = InStr(p + Len(vgPeloQue$), x$, vgOQue$)                     'na string alvo
        Wend
    End If
    Substitui$ = x$                               'retorna a nova string
End Function



'parseia string St$ atraves do caracter Delim$
Public Function Parse(ByRef St As Variant, ByVal Delim As String, Optional ByVal NumParse As Integer = 0) As String
    Dim i As Integer, NewSt As String, RetVal As String, Cont As Integer
    NewSt$ = St
PegaOutro:
    Cont = Cont + 1
    i = InStr(NewSt$, Delim$)
    If i > 0 Then
        RetVal$ = Left$(NewSt$, i - 1)
        NewSt$ = Mid$(NewSt$, i + Len(Delim$))
        If Cont < NumParse Then
            GoTo PegaOutro
        End If
    Else
        If NumParse = 0 Or Cont = NumParse Then
            RetVal$ = NewSt$
            NewSt$ = ""
        Else
            RetVal$ = ""
        End If
    End If
    If NumParse = 0 Then
        St = NewSt$
    End If
    Parse$ = RetVal$
End Function

'verifica se variável/campo esta vazio
Public Function Vazio(ByVal vgSt As Variant) As Integer
 If IsNull(vgSt) Or IsEmpty(vgSt) Then            'se está nulo ou vazio
  Vazio = True                                    'retorna sim
 Else
  Select Case VarType(vgSt)                       'tipo do campo/variável
   Case 8                                         'string
    Vazio = (Len(Trim$(vgSt)) = 0)                'se o tamanho é zero
   Case 7                                         'data
    Vazio = (vgSt <= CDate("2/1/100"))            'menor que 2/1/100
   Case Else                                      'numérico/logico
    Vazio = (vgSt = 0)                            'se for igual a zero
  End Select
 End If
End Function

Public Function FormatoDecimal(AValor As Double) As String
    Dim v As String
    v = CStr(AValor)
    If InStr(1, v, ",") > 0 Then
        Mid$(v, InStr(1, v, ","), 1) = "."
    End If
    FormatoDecimal = v
End Function



'Verifica se arquivo existe
'-1 = Arquivo existe
' 0 = Arquivo não existe
' 2 = Erro! Não existe, diretório inválido ou compartilhado ou Drive não preparado
Public Function Existe(ByVal Arq As String) As Integer
 On Error Resume Next
 If Len(Arq) > 0 Then
  Existe = (Len(Dir$(Arq$, vbArchive Or vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)) > 0)
  If Err Then
   Err.Clear
   Existe = 2
  End If
 Else
  Existe = 2
 End If
End Function

'pega posição de um argumento, a partir da direita
Public Function Rat(vgAlvo As String, Pesq As String) As Integer
 Dim i As Integer, RetVal As Integer, _
     j As Integer, p As String                    'dimensiona
 RetVal = False                                   'prepara retorno falso
 j = Len(Pesq$)                                   'pega tamanho da string a pesquisar
 For i = Len(vgAlvo$) To 1 Step -1                'corre de trás para a frente
  p$ = Mid$(vgAlvo$, i, j)                        'para comparar...
  If p$ = Pesq$ Then                              'se achou,
   RetVal = i                                     'prepara para retornar posição
   Exit For                                       'sai do for...
  End If
 Next
 Rat = RetVal                                     'retorna posição
End Function

'enche caracteres à direita de uma string
Public Function RPad(vgSt As Variant, vgTm As Integer, vgCh As String) As String
 Dim x As String                                            'dimensiona
 If VarType(vgSt) = vbString Then                           'se veio uma string
  x$ = vgSt                                                 'pega ela...
 Else                                             'senão,
  x$ = CStr(vgSt)                                 'transforma em string
 End If
 RPad$ = Left$(LTrim$(x$) + String$(vgTm, vgCh$), vgTm) 'completa com brancos à direita
End Function

'enche caracteres à esquerda de uma string
Public Function LPad(vgSt As Variant, vgTm As Integer, vgCh As String) As String
 Dim x As String                                            'dimensiona
 If VarType(vgSt) = vbString Then                           'se veio uma string
  x$ = vgSt                                                 'pega ela...
 Else                                             'senão,
  x$ = CStr(vgSt)                                 'transforma em string
 End If
 LPad$ = Right$(String$(vgTm, vgCh$) + LTrim$(x$), vgTm) 'completa com brancos à esquerda
End Function

'Executa o comando SQL no banco de dados com retorno
Public Function SQLExecutaRetorno(ComandoSQL As String, Campo As String, Optional Padrao As Variant) As Variant
'Executa o camando SQL
On Error GoTo deuErro
SQLExecutaRetorno = vgDb.Execute(ComandoSQL)(Campo)
If IsNull(SQLExecutaRetorno) Then SQLExecutaRetorno = Padrao
Exit Function

deuErro:
 Err.Clear
 SQLExecutaRetorno = Padrao
End Function

Public Function fSQL(Campo As Variant, Optional QtDecimais As Integer = 4) As String
   If QtDecimais < 1 Or QtDecimais > 20 Then QtDecimais = 4
   fSQL = Substitui(Format(Campo, "######0." & LPad("", QtDecimais, "0")), ",", ".", UM_A_UM)
End Function

Public Function FSQL1(Campo As String) As String
   FSQL1 = Substitui(Campo, "'", Chr(34), UM_A_UM)
End Function

'Formata o campo data para ser usado no SQL
Public Function FdtSQL(DATA As Variant) As String
 FdtSQL = "'" & Format(DATA, "yyyymmdd") & "'"
End Function

'Formata o campo data para ser usado no SQL
Public Function FhrSQL(hora As Variant) As String
 FhrSQL = "'1899-01-01 " & Format(hora, "hh:mm") & "'"
End Function

'Formata o campo data/hora para ser usado no SQL
Public Function FdthrSQL(DataHora As Variant) As String
 FdthrSQL = "'" & Format(DataHora, "yyyymmdd hh:MM:ss") & "'"
End Function

'Verifica se o programa está sendo executado via VB
Public Function VbInDesign() As Boolean
 On Error GoTo deuErro
 Debug.Print 1 / 0
 VbInDesign = False
 Exit Function

deuErro:
 VbInDesign = True
End Function

