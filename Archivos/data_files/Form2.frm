VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Descompactar"
      Height          =   1005
      Left            =   360
      TabIndex        =   1
      Top             =   1665
      Width           =   3885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   780
      Left            =   405
      TabIndex        =   0
      Top             =   405
      Width           =   4020
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form3.Show
Unload Me
End Sub
