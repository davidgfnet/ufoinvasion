VERSION 5.00
Begin VB.Form Juego 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "UFO Invasion"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   Icon            =   "Juego.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Juego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyUp
    KArriba = True
Case vbKeyDown
    KAbajo = True
Case vbKeyLeft
    KIzquierda = True
Case vbKeyRight
    KDerecha = True
Case vbKeyEscape
    KEscape = True
Case vbKeyReturn
    KIntro = True
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyUp
    KArriba = False
Case vbKeyDown
    KAbajo = False
Case vbKeyLeft
    KIzquierda = False
Case vbKeyRight
    KDerecha = False
Case vbKeyEscape
    KEscape = False
Case vbKeyReturn
    KIntro = False
Case vbKeyPause
    JPausado = Not JPausado
End Select
End Sub
