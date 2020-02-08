Attribute VB_Name = "SndEngine"
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'--------------------------------------------------------------------------------------------
'======================================== SOUND ENGINE ======================================
'--------------------------------------------------------------------------------------------
Private Function OpenMultimedia(hwnd As Long, AliasName As String, filename As String, typeDevice As String) As String
Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128
Dim tmp As String * 255
Dim lenShort As Long
Dim ShortPathAndFile As String
Const WS_CHILD = &H40000000

lenShort = GetShortPathName(filename, tmp, 255)
ShortPathAndFile = Left$(tmp, lenShort)

cmdToDo = "open " & ShortPathAndFile & " type " & typeDevice & " Alias " & AliasName & " parent " & hwnd & " Style " & WS_CHILD
dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)

OpenMultimedia = "1"
End Function

Private Function PlayMultimedia(AliasName As String, from_where As String, to_where As String) As String
If from_where = vbNullString Then from_where = 0
If to_where = vbNullString Then to_where = GetTotalframes(AliasName)

If AliasName = glo_AliasName Then
    glo_from = from_where
    glo_to = to_where
End If

Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128

cmdToDo = "play " & AliasName & " from " & from_where & " to " & to_where

dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)

PlayMultimedia = "1"
End Function

Private Function GetTotalframes(AliasName As String) As Long
Dim dwReturn As Long
Dim Total As String * 128

dwReturn = mciSendString("set " & AliasName & " time format frames", Total, 128, 0&)
dwReturn = mciSendString("status " & AliasName & " length", Total, 128, 0&)

If Not dwReturn = 0 Then
    GetTotalframes = -1
    Exit Function
End If
GetTotalframes = Val(Total)
End Function

Private Function StopMultimedia(AliasName As String) As String
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Stop " & AliasName, 0&, 0&, 0&)
StopMultimedia = "1"
End Function

Private Function PauseMultimedia(AliasName As String) As String
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Pause " & AliasName, 0&, 0&, 0&)
PauseMultimedia = "1"
End Function

Private Function ResumeMultimedia(AliasName As String) As String
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Resume " & AliasName, 0&, 0&, 0&)
ResumeMultimedia = "1"
End Function


Private Function SetVolume(AliasName As String, Channel As String, VolumeValue As Long) As String
Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128
Dim VolumeV As Long
VolumeV = VolumeValue

If VolumeV < 0 Or VolumeV > 100 Then
    SetVolume = "out of volume"
    Exit Function
End If

VolumeV = VolumeV * 10

If LCase(Channel) = "left" Or LCase(Channel) = "right" Then
    cmdToDo = "setaudio " & AliasName & " " & Channel & " Volume to " & VolumeV
Else
    cmdToDo = "setaudio " & AliasName & " Volume to " & VolumeV
End If

dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
SetVolume = "Success"
End Function

Private Function GetVolume(AliasName As String, Channel As String) As Long
Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim Volume As String * 128

If LCase(Channel) = "left" Or LCase(Channel) = "right" Then
    cmdToDo = "status " & AliasName & " " & Channel & " Volume"
Else
    cmdToDo = "status " & AliasName & " Volume"
End If

dwReturn = mciSendString(cmdToDo, Volume, 128, 0&)
GetVolume = Val(Volume) / 10
End Function

Private Function CloseAll() As String
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Close All", 0&, 0&, 0&)
CloseAll = "Success"
End Function

Public Function GetStatusMultimedia(AliasName As String) As String
Dim dwReturn As Long
Dim status As String * 128
Dim ret As String * 128

dwReturn = mciSendString("status " & AliasName & " mode", status, 128, 0&)  'Get status

If Not dwReturn = 0 Then  'not success
    GetStatusMultimedia = "ERROR"
    Exit Function
End If

'Extract just the string
Dim i As Integer
Dim CharA As String
Dim RChar As String
RChar = Right$(status, 1)
For i = 1 To Len(status)
    CharA = Mid(status, i, 1)
    If CharA = RChar Then Exit For
    GetStatusMultimedia = GetStatusMultimedia + CharA
Next i
End Function

'--------------------------------------------------------------------------------------------
'======================================== DAVID =============================================
'--------------------------------------------------------------------------------------------
Public Sub CargarSonidoWin(ByVal File As String, ByVal Alias As String, ByVal hwnd As Long)
Dim ret As String
ret = OpenMultimedia(hwnd, Alias, File, "MPEGVideo")
If ret <> "1" Then
    MsgBox "Archivo no cargado"
End If
End Sub

Public Sub ReproducirSonidoWin(ByVal Alias As String)
Call PlayMultimedia(Alias, vbNullString, vbNullString)
End Sub

Public Sub PararSonidoWin(ByVal Alias As String)
StopMultimedia Alias
End Sub

Public Sub PauseSonidoWin(ByVal Alias As String)
PauseMultimedia Alias
End Sub

Public Sub ReanudarSonidoWin(ByVal Alias As String)
ResumeMultimedia Alias
End Sub

Public Sub SetVolumenWin(ByVal Alias As String, ByVal Volumen As Long)
If Volumen >= 0 And Volumen <= 100 Then
    SetVolume Alias, "all", Volumen
End If
End Sub

Public Function GetVolumenWin(ByVal Alias As String) As String
GetVolumen = GetVolume(Alias, "all")
End Function

Public Sub CerrarTodosSonidosWin()
CloseAll
End Sub

Public Function Parado(ByVal Alias As String) As Boolean
If GetStatusMultimedia(Alias) = "stopped" Then Parado = True
End Function
