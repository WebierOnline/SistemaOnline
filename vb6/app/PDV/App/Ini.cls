VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileStruct Lib "kernel32" Alias "GetPrivateProfileStructA" (ByVal lpszSection As String, ByVal lpszKey As String, lpStruct As Byte, ByVal uSizeStruct As Long, ByVal szFile As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileStruct Lib "kernel32" Alias "WritePrivateProfileStructA" (ByVal lpszSection As String, ByVal lpszKey As String, lpStruct As Byte, ByVal uSizeStruct As Long, ByVal szFile As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private mArqvIni As String

Public Property Get Arquivo() As String
   Arquivo = mArqvIni
End Property

Public Property Let Arquivo(ByVal NewFile As String)
   mArqvIni = NewFile
End Property

Public Function LerTexto(Secao As String, Chave As String, Optional Padrao As String = "") As String
   Dim BufferSize As Long, Ret As Long
   
   'Increase the buffer until it's big enough
   Do
      BufferSize = BufferSize + 256
      LerTexto = Space(BufferSize)
      Ret = GetPrivateProfileString(Secao, Chave, Padrao, LerTexto, BufferSize, mArqvIni)
   Loop Until Ret < BufferSize - 1
   
   'Chop off the trailing Chr$(0)'s
   If Ret > 0 Then
      LerTexto = Left$(LerTexto, Ret)
    Else
      LerTexto = ""
   End If
End Function

Public Sub EscreverTexto(Secao As String, Chave As String, Valor As String)
   Dim l As Long
   WritePrivateProfileString Secao, Chave, Valor, mArqvIni
End Sub

Public Sub ExcluirChave(Secao As String, Chave As String)
   WritePrivateProfileString Secao, Chave, "", mArqvIni
End Sub

Public Function LerSecao(Secao As String) As Collection
   Dim BufferSize As Long, Ret As Long, Cnt As Long
   Dim Buffer As String, Splitted() As String
   
   Set LerSecao = New Collection
   
   'Increase the buffer until it's big enough
   Do
      BufferSize = BufferSize + 1024
      Buffer = Space(BufferSize)
      Ret = GetPrivateProfileSection(Secao, Buffer, BufferSize, mArqvIni)
   Loop Until Ret < BufferSize - 2
   
   If Ret > 0 Then
      'Split the string
      Ret = SplitString(Splitted(), Left$(Buffer, Ret - 1))
      
      'Add each part to the collection
      For Cnt = 0 To Ret - 1
         LerSecao.Add Splitted(Cnt)
      Next
   End If
End Function

Public Function LerNomesSecao(Secao As String) As Collection
   Dim BufferSize As Long, Ret As Long, Cnt As Long
   Dim Buffer As String, Splitted() As String
   
   Set LerNomesSecao = New Collection
   
   'Increase the buffer until it's big enough
   Do
      BufferSize = BufferSize + 1024
      Buffer = Space(BufferSize)
      Ret = GetPrivateProfileSectionNames(Buffer, BufferSize, mArqvIni)
   Loop Until Ret < BufferSize - 2
   
   If Ret > 0 Then
      'Split the string
      Ret = SplitString(Splitted(), Left$(Buffer, Ret - 1))
      'Add each part to the collection
      For Cnt = 0 To Ret - 1
         LerNomesSecao.Add Splitted(Cnt)
      Next
   End If
End Function

'We don't use VB6's Split method, to maintain VB5 compatibility
Private Function SplitString(sStringArray() As String, ByVal sInput As String) As Long
   Dim ZeroPos As Long, ArrayCount As Long
   
   ZeroPos = InStr(1, sInput, Chr$(0), vbBinaryCompare)
   While ZeroPos > 0
      ReDim Preserve sStringArray(0 To ArrayCount) As String
      sStringArray(ArrayCount) = Left$(sInput, ZeroPos - 1)
      sInput = Right$(sInput, Len(sInput) - ZeroPos)
      ArrayCount = ArrayCount + 1
      ZeroPos = InStr(1, sInput, Chr$(0), vbBinaryCompare)
   Wend
   
   If sInput <> "" Then
      ReDim Preserve sStringArray(0 To ArrayCount) As String
      sStringArray(ArrayCount) = sInput
      ArrayCount = ArrayCount + 1
   End If
   
   SplitString = ArrayCount
End Function
