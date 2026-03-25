VERSION 5.00
Begin VB.Form ListarPDV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de terminais"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "ListarPDV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mExec As Boolean

Public Property Get Done() As Boolean
   Done = mExec
End Property

Sub pListarPDVs()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT * FROM pdvs WHERE (ativo = 1);"
   Set r = dbData.OpenRecordset(sSQL)
   
   List1.Clear
   
   Do While Not r.EOF
      List1.AddItem r("descricao")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If List1.ListCount = 0 Then
      ShowMsg "Esta operação não pode ser realizada.", vbExclamation
      Exit Sub
   End If
   
   Dim fLib As LiberarVenda
   Dim bCancel As Boolean
   
   Set fLib = New LiberarVenda
   Load fLib
   
   fLib.Show vbModal
   bCancel = fLib.Cancelled
   
   Unload fLib
   Set fLib = Nothing
   
   If bCancel Then Exit Sub
   
   dbData.Execute "UPDATE pedidos SET status_pedido = -1, data_compra = CONVERT(DATETIME, '" & Format$(Now, ocDATA) & "', 103), maquina = '" & List1.Text & "' WHERE (cod_pedido = " & PDV.txtCodPedido & ");"
   ShowMsg "Transferência realizada com sucesso!" & vbCr & vbCr & "Encaminha o cliente para o caixa '" & List1.Text & "'.", vbInformation
   mExec = True
   Unload Me
End Sub

Private Sub Form_Load()
   CenterForm Me, Width, Height
   pListarPDVs
   mExec = False
End Sub
