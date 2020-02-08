VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   1065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form3"
   ScaleHeight     =   1065
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Descompactar fichero"
      Height          =   420
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   3705
   End
   Begin MSComDlg.CommonDialog Cuadro 
      Left            =   315
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cuadro.DialogTitle = "Descompactar archivo"
Cuadro.CancelError = True
Cuadro.FileName = ""
Cuadro.Filter = "Archivos Data|*.dat"
Cuadro.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNExplorer
Cuadro.MaxFileSize = 32000

On Local Error Resume Next
Cuadro.ShowOpen

If Err.Number <> 0 Then Exit Sub

Dim libre As Integer, int1 As Integer, lng1 As Long, str1 As String
libre = FreeFile
Open Cuadro.FileName For Binary As #libre
str1 = Space(19)
Get #libre, , str1
10
Get #libre, , int1
str1 = Space(int1)
Get #libre, , str1
If str1 = "ENDOFFILE" Then GoTo 20
Open App.Path & "\" & str1 For Output As #5
Get #libre, , lng1
str1 = Space(lng1)
Get #libre, , str1
Print #5, str1;
Close #5
GoTo 10
20
Close
End Sub

