VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   180
   ClientTop       =   795
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleMode       =   0  'User
   ScaleWidth      =   6885
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   66
      Left            =   79
      TabIndex        =   0
      Top             =   79
      Width           =   66
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

    DataType.Initialize

    Me.Label1.Caption = "Tutorial 04" & vbCrLf & "---------------" & vbCrLf

    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")

    ' Create two sheets
    workbook.easy_addWorksheet_2 ("First tab")
    workbook.easy_addWorksheet_2 ("Second tab")

    ' Get the table of data for the first worksheet
    Set xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()

    ' Add data in cells for report header
    For Column = 0 To 4
        xlsFirstTable.easy_getCell(0, Column).setValue ("Column " & (Column + 1))
        xlsFirstTable.easy_getCell(0, Column).setDataType (DataType.DATATYPE_STRING)
    Next

    ' Add data in cells for report values
    For Row = 0 To 99
        For Column = 0 To 4
            xlsFirstTable.easy_getCell(Row + 1, Column).setValue ("Data " & (Row + 1) & ", " & (Column + 1))
            xlsFirstTable.easy_getCell(Row + 1, Column).setDataType (DataType.DATATYPE_STRING)
        Next
    Next

    ' Export the XLSX file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & _
                           "Writing file C:\Samples\Tutorial04 - export data to Excel.xlsx"
    workbook.easy_WriteXLSXFile "C:\Samples\Tutorial04 - export data to Excel.xlsx"

    ' Confirm export of Excel file
    If workbook.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
    End If

    ' Dispose memory
    workbook.Dispose
End Sub



