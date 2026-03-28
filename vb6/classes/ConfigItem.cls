VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mName As String
Private mValue As Variant

Friend Sub SetValues(ByVal sName As String, ByVal vValue As Variant)
   mName = sName
   mValue = vValue
End Sub

Public Property Get Name() As String
   Name = mName
End Property

Public Property Let Name(ByVal NewValue As String)
   mName = NewValue
End Property

Public Property Get Value() As Variant
   Value = mValue
End Property

Public Property Let Value(ByVal NewValue As Variant)
   mValue = NewValue
End Property
