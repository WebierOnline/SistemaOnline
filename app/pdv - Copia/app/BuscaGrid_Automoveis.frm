VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form BuscaGrid_Automoveis 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstBusca 
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   5847
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "BuscaGrid_Automoveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mCancelled As Boolean
Dim vInfo() As String

Public Property Get Cancelled() As Boolean
   Cancelled = mCancelled
End Property

Public Property Get InfoProduct() As String()
   InfoProduct = vInfo
End Property

Sub pCriarGrid()
   lstBusca.FullRowSelect = True
   lstBusca.LabelEdit = lvwManual
   lstBusca.Visible = True
   lstBusca.View = lvwReport
   lstBusca.HideSelection = False
   lstBusca.ListItems.Clear
   
   lstBusca.ColumnHeaders.Clear
   lstBusca.ColumnHeaders.Add , , "CÓDIGO", 0
   lstBusca.ColumnHeaders.Add , , "COD_BARRA", 0
   lstBusca.ColumnHeaders.Add , , "DESCRIÇÃO", 3800
   lstBusca.ColumnHeaders.Add , , "QTDE", 650, 1
   lstBusca.ColumnHeaders.Add , , "COMPATIBILIDADE", 4000, 0
   lstBusca.ColumnHeaders.Add , , "LOC.", 650, 1
   lstBusca.ColumnHeaders.Add , , "VALOR", 800, 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 0 Then
      If KeyCode = vbKeyEscape Then Unload Me
      If KeyCode = vbKeyReturn Then lstBusca_KeyDown KeyCode, Shift
   End If
End Sub

Private Sub Form_Load()
   Set Icon = Nothing
   KeyPreview = True
   mCancelled = True
   Erase vInfo
   pCriarGrid
End Sub

Private Sub lstBusca_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If Shift = 0 And KeyCode = vbKeyReturn Then
      If lstBusca.ListItems.Count = 0 Then
         ShowMsg "Nenhum item disponível para seleção.", vbExclamation
         Exit Sub
      End If
      
      If lstBusca.SelectedItem Is Nothing Then
         ShowMsg "Nenhum item foi selecionado.", vbExclamation
         Exit Sub
      End If
      
      If Not lstBusca.SelectedItem.Selected Then
         ShowMsg "Nenhum item foi selecionado.", vbExclamation
      End If
      
      ReDim vInfo(1 To 4)
      vInfo(1) = lstBusca.SelectedItem
      vInfo(2) = lstBusca.SelectedItem.ListSubItems.Item(1).Text
      vInfo(3) = lstBusca.SelectedItem.ListSubItems.Item(2).Text
      vInfo(4) = lstBusca.SelectedItem.ListSubItems.Item(3).Text
      
      mCancelled = False
      Unload Me
   End If

End Sub
