Attribute VB_Name = "Module1"
Public Function LeerA(ByVal Longitud As Double, ByVal Camino As Integer, Optional Barra As Object) As String
On Local Error Resume Next
'de 200000 en 200000
DoEvents
If Longitud <= 200000 Then
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
    If a3 >= 200000 Then
        Strin = Space(200000)
        Get #Camino, , Strin
        LeerA = LeerA & Strin
    Else
        Strin = Space(a3)
        Get #Camino, , Strin
        LeerA = LeerA & Strin
    End If
    Barra.Value = a2
    a3 = a3 - 200000
    DoEvents
Next
Barra.Value = 0
End Function

