VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ReinicarVenda 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwPed 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3413
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "ReinicarVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCancelled As Boolean
Private mNroPedido As Long

Public Property Get Cancelled() As Boolean
   Cancelled = mCancelled
End Property

Public Property Get OrderNumber() As Long
   OrderNumber = mNroPedido
End Property

Sub pCriarGrade()
   With lvwPed
      .ColumnHeaders.Clear
      .ColumnHeaders.Add , , "N.º pedido", 1200
      .ColumnHeaders.Add , , "Data", 1200, 1
      .ColumnHeaders.Add , , "Valor", 1200, 2
      
      .View = lvwReport
      .FullRowSelect = True
      .HideSelection = False
   End With
End Sub

Sub pPositionCtrls()
   cmdCancelar.Move ScaleWidth - cmdCancelar.Width - 120, ScaleHeight - cmdCancelar.Height - 120
   cmdOK.Move cmdCancelar.Left - cmdOK.Width - 120, cmdCancelar.Top
   lvwPed.Move 0, 0, ScaleWidth, ScaleHeight - cmdOK.Height - 240
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   
   If lvwPed.ListItems.Count = 0 Then
      ShowMsg "Não existem pedidos para serem reiniciados.", vbExclamation
      Exit Sub
   End If
   
   If lvwPed.SelectedItem Is Nothing Then
      ShowMsg "Selecione um pedido para continuar.", vbExclamation
      Exit Sub
   End If
   
   If Not lvwPed.SelectedItem.Selected Then
      ShowMsg "Selecione um pedido para continuar.", vbExclamation
      Exit Sub
   End If
   
   mNroPedido = lvwPed.SelectedItem.Text
   mCancelled = False
   Unload Me
   
End Sub

Private Sub Form_Load()
   pCriarGrade
   mCancelled = True
   mNroPedido = 0
End Sub

Private Sub Form_Resize()
   If WindowState <> vbMinimized Then pPositionCtrls
End Sub
