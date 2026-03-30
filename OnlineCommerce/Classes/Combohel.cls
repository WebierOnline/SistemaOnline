VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComboHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'usado no Projeto OnlineCommerce
Option Explicit

Private WithEvents moCombo As ComboBox
Attribute moCombo.VB_VarHelpID = -1

Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const CB_SETCURSEL = &H14E
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158

Private Function MatchingRow(ByVal sText$, ByVal mStart&, ByVal iMatchType%) As Long
   Dim mFoundRow As Long                          'the row we find
   
   On Error GoTo Err_MatchingRow
   mFoundRow = SendMessageByString(moCombo.hwnd, iMatchType, mStart, sText)
   
Bye_MatchingRow:
   MatchingRow = mFoundRow
   Exit Function

Err_MatchingRow:
   mFoundRow = -1
   Resume Bye_MatchingRow
End Function

Public Sub AttachTo(ByVal oCombo As ComboBox)
   On Error Resume Next
   Set moCombo = oCombo
End Sub

Private Sub Class_Terminate()
   On Error Resume Next
   Set moCombo = Nothing
End Sub

Public Function Contains(ByVal sText$) As Boolean
   On Error Resume Next
   Contains = (SendMessageByString(moCombo.hwnd, CB_FINDSTRINGEXACT, 0, sText) <> -1)
End Function

Private Function ResultingText(iKeyAscii%) As String
   Dim sLeft As String                                     'string element
   Dim sSel As String                                      'selected string element
   Dim sRight As String                                    'string element
   Dim sResult As String                                   'what we'll return
   
   On Error Resume Next
   
   With moCombo
      sLeft = Left$(.Text, .SelStart)                       'SelStart is 0-based
      sSel = Mid$(.Text, .SelStart + 1, .SelLength)
      sRight = Mid$(.Text, .SelStart + .SelLength + 1)
   End With
   
   Select Case iKeyAscii
      Case vbKeyBack                                        'Backspace key
         If Len(sSel) = 0 Then                              'nothing selected
            sResult = MinusRightChar(sLeft) & sRight        'delete first char on the left
         Else                                               'selection exists
            sResult = sLeft & sRight                        'delete selected text only
         End If
      
      Case vbKeyDelete                                      'Delete key
         If Len(sSel) = 0 Then                              'nothing selected
            sResult = sLeft & MinusLeftChar(sRight)         'delete first char on the right
         Else                                               'selection exists
            sResult = sLeft & sRight                        'delete selected text only
         End If
      
      Case Else                                             'an ordinary character
         sResult = sLeft & Chr$(iKeyAscii) & sRight
      
   End Select
   
Bye_ResultingText:
   ResultingText = sResult
End Function

Private Function MinusLeftChar(ByVal sGiven As String) As String
   On Error Resume Next
   
   If Len(sGiven) = 0 Then
      MinusLeftChar = ""
   Else
      MinusLeftChar = Mid$(sGiven, 2)
   End If
End Function

Private Function MinusRightChar(ByVal sGiven As String) As String
   On Error Resume Next
   
   If Len(sGiven) = 0 Then
      MinusRightChar = ""
   Else
      MinusRightChar = Left$(sGiven, Len(sGiven) - 1)
   End If
End Function

Private Sub moCombo_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim sSearchOn As String                             'current string to search on
   
   On Error Resume Next
   If KeyCode <> vbKeyDelete Then GoTo Bye_moCombo_KeyDown
   sSearchOn = ResultingText(KeyCode)
   If SearchOn(sSearchOn, CB_FINDSTRINGEXACT) = True Then KeyCode = 0
   
Bye_moCombo_KeyDown:
   Exit Sub
End Sub

Private Sub moCombo_KeyPress(KeyAscii As Integer)
   Dim sSearchOn As String                                 'current string to search on
   
   On Error Resume Next
   
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
      GoTo Bye_moCombo_KeyPress                             'we don't want to know
   End If
   
   sSearchOn = ResultingText(KeyAscii)
   
   If KeyAscii = vbKeyBack Then                            'prevent re-finding original !
      If SearchOn(sSearchOn, CB_FINDSTRINGEXACT) = True Then KeyAscii = 0
   Else
      If SearchOn(sSearchOn, CB_FINDSTRING) = True Then KeyAscii = 0
   End If
   
Bye_moCombo_KeyPress:
   Exit Sub
End Sub

Private Function SearchOn(ByVal sStartText$, ByVal iMatchType%) As Boolean
   Dim mOriginalIndex As Long
   Dim mNewIndex As Long
   Dim yMoveOccurred As Boolean
   
   On Error Resume Next
   mOriginalIndex = moCombo.ListIndex
   yMoveOccurred = False
   
   mOriginalIndex = mOriginalIndex - 1
   If mOriginalIndex < -1 Then mOriginalIndex = -1
   
   With moCombo
      If Len(sStartText) > 0 Then                           'if it contains any text
         mNewIndex = MatchingRow(sStartText, mOriginalIndex, iMatchType)
         If mNewIndex <> -1 Then                             'match found
            SendMessageByString .hwnd, CB_SETCURSEL, mNewIndex, 0    'set list index
            .SelStart = Len(sStartText)
            .SelLength = Len(.Text) - Len(sStartText)         'select the bit AFTER the match
            yMoveOccurred = True
         End If
      Else                                                  'contains no text
         SendMessageByString .hwnd, CB_SETCURSEL, -1, 0      'set list index to -1
         If mOriginalIndex <> -1 Then yMoveOccurred = True
      End If
   End With
   
   SearchOn = yMoveOccurred
End Function
