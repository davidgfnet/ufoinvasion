Attribute VB_Name = "Datafiles"
Option Explicit
Dim a1 As Long, a2 As Long, a3 As Long
Public Function LeerA(ByVal Longitud As Double, ByVal Camino As Integer, Optional Barra As Object) As String
Dim CadenaRapida As StringConcat, Strin As String
Set CadenaRapida = New StringConcat
On Local Error Resume Next
'de 200000 en 200000
DoEvents
If Longitud <= 200000 Then
    Cargador.Barra.Width = (Cargador.BarraMax.Width / LOF(Camino)) * Seek(Camino)
    LeerA = Space(Longitud)
    Get #Camino, , LeerA
    Exit Function
End If
If Longitud Mod 200000 = 0 Then
    a1 = Longitud / 200000
Else
    a1 = Fix(Longitud / 200000) + 1
End If

Barra.Value = 0
Barra.Min = 0
Barra.Max = a1
a3 = Longitud
For a2 = 1 To a1
    Cargador.Barra.Width = (Cargador.BarraMax.Width / LOF(Camino)) * Seek(Camino)
    If a3 >= 200000 Then
        Strin = Space(200000)
        Get #Camino, , Strin
        CadenaRapida.Add Strin
    Else
        Strin = Space(a3)
        Get #Camino, , Strin
        CadenaRapida.Add Strin
    End If
    Barra.Value = a2
    a3 = a3 - 200000
    DoEvents
Next
Barra.Value = 0
LeerA = CadenaRapida.GetStr()
End Function

Public Sub DescompactarFichero(ByVal Fic As String, Carpeta As String)
Dim libre As Integer, int1 As Integer, lng1 As Long, str1 As String, libre2 As Integer
If Right(Carpeta, 1) <> "\" Then Carpeta = Carpeta & "\"
libre = FreeFile
Open Fic For Binary As #libre
str1 = Space(19)
Get #libre, , str1
10
Get #libre, , int1
str1 = Space(int1)
Get #libre, , str1
If str1 = "ENDOFFILE" Then GoTo 20
libre2 = FreeFile
CrearCarpeta Carpeta
Open Carpeta & str1 For Output As #libre2
Get #libre, , lng1
Print #libre2, LeerA(lng1, libre);
Close #libre2
GoTo 10
20
Close
End Sub

Public Sub CrearCarpeta(ByVal Carpeta As String)
On Local Error Resume Next
MkDir Carpeta
End Sub

Public Function Existe(ByVal Fic As String) As Boolean
On Local Error Resume Next
Err.Clear: Err.Number = 0
Dim nada As VbFileAttribute
nada = GetAttr(Fic)
If Err.Number = 0 Then Existe = True
Err.Clear: Err.Number = 0
End Function
