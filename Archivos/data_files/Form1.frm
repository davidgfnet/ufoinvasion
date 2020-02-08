VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Crear fichero"
      Height          =   375
      Left            =   6210
      TabIndex        =   2
      Top             =   6840
      Width           =   2985
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Añadir"
      Height          =   375
      Left            =   225
      TabIndex        =   1
      Top             =   6840
      Width           =   3120
   End
   Begin VB.ListBox Lista 
      Height          =   6495
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   8925
   End
   Begin MSComDlg.CommonDialog Cuadro 
      Left            =   8145
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cuadro.DialogTitle = "Añadir archivo"
Cuadro.CancelError = True
Cuadro.FileName = ""
Cuadro.Filter = "Todos los archivos|*.*"
Cuadro.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer
Cuadro.MaxFileSize = 32000

On Local Error Resume Next
Cuadro.ShowOpen

If Err.Number <> 0 Then Exit Sub

Dim tss As String
Dim tss2 As String
Dim i As Integer, j As Integer, d As Integer
Dim Matriz()

If InStr(1, Cuadro.FileName, Chr(0)) = 0 Then
    ReDim Matriz(1 To 1)
    Matriz(1) = Cuadro.FileName
    d = 1
    tss2 = APath(Cuadro.FileName)
    GoTo 4
End If

Dim ant As Long
For i = Len(Cuadro.FileName) To 1 Step -1
    DoEvents
    If Mid(Cuadro.FileName, i, 1) = "\" Then
        tss = Left(Cuadro.FileName, i - 1)
        d = i
        Exit For
    End If
Next
For i = d To Len(Cuadro.FileName)
    DoEvents
    If Mid(Cuadro.FileName, i, 1) = Chr(0) Then
        tss2 = tss & Mid(Cuadro.FileName, d, i - d)
        Exit For
    End If
Next
If Right(tss2, 1) <> "\" Then tss2 = tss2 & "\"
ant = i + 1
For d = (i + 1) To Len(Cuadro.FileName)
    DoEvents
    If Mid(Cuadro.FileName, d, 1) = Chr(0) Or d = Len(Cuadro.FileName) Then
        j = j + 1
        ReDim Preserve Matriz(1 To j)
        If d <> Len(Cuadro.FileName) Then
            Matriz(j) = tss2 & Mid(Cuadro.FileName, ant, d - ant)
        Else
            Matriz(j) = tss2 & Mid(Cuadro.FileName, ant)
        End If
        ant = d + 1
    End If
Next
d = j

4
For j = 1 To d
    If NombreR(Matriz(j)) = "" Then GoTo 30
    DoEvents
    If FileLen(Matriz(j)) = 0 Then GoTo 30
    Lista.AddItem Matriz(j)
30
Next
End Sub

Private Function APath(ByVal Nombre As String) As String
If Nombre = "" Then Exit Function
For xxx = Len(Nombre) To 1 Step -1
    xx = InStr(xxx, Nombre, "\")
    If xx <> 0 Then Exit For
Next
APath = Mid$(Nombre, 1, xx - 1)
End Function

Private Function NombreR(ByVal Nombre As String) As String
If Nombre = "" Then Exit Function
For xxx = Len(Nombre) To 1 Step -1
    xx = InStr(xxx, Nombre, ".")
    If xx <> 0 Then Exit For
Next
NombreR = Mid$(Nombre, xx + 1, Len(Nombre) - xx + 1)
End Function

Private Function NombreArchivo(ByVal Nombre As String) As String
If Nombre = "" Then Exit Function
For xxx = Len(Nombre) To 1 Step -1
    xx = InStr(xxx, Nombre, "\")
    If xx <> 0 Then Exit For
Next
NombreArchivo = Mid$(Nombre, xx + 1, Len(Nombre) - xx + 1)
End Function

Private Sub Command2_Click()
Dim libre As Integer, str1 As String, libre2 As Integer, lng1 As Long
libre = FreeFile
Open App.Path & "\Game.dat" For Binary As #libre
Put #libre, , "UFOINVASIONGAMEFILE"
Dim int1 As Integer, int2 As Integer
For int2 = 0 To Lista.ListCount - 1
    int1 = Len(NombreArchivo(Lista.List(int2)))
    Put #libre, , int1
    Put #libre, , NombreArchivo(Lista.List(int2))
    lng1 = FileLen(Lista.List(int2))
    Put #libre, , lng1
    libre2 = FreeFile
    Open Lista.List(int2) For Binary As #libre2
    str1 = Space(LOF(libre2))
    Get #libre2, , str1
    Close #libre2
    Put #libre, , str1
Next
int1 = Len("ENDOFFILE")
Put #libre, , int1
Put #libre, , "ENDOFFILE"
Close
End Sub

