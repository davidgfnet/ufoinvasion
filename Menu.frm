VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UFO INVASION - BETA VERSION"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Menu.frx":038A
   ScaleHeight     =   4845
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image salir 
      Height          =   750
      Left            =   1245
      Picture         =   "Menu.frx":7E7E
      Top             =   3060
      Width           =   4965
   End
   Begin VB.Image start 
      Height          =   750
      Left            =   1245
      Picture         =   "Menu.frx":9510
      Top             =   990
      Width           =   4965
   End
   Begin VB.Image ins 
      Height          =   750
      Left            =   1245
      Picture         =   "Menu.frx":BDEC
      Top             =   2025
      Width           =   4965
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Seleccion = 0
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub start_Click()
Seleccion = 1
Unload Me
End Sub
