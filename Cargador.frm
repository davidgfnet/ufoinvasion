VERSION 5.00
Begin VB.Form Cargador 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BarraMax 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   180
      ScaleHeight     =   105
      ScaleWidth      =   5700
      TabIndex        =   0
      Top             =   4590
      Width           =   5730
      Begin VB.Label Barra 
         BackColor       =   &H00009600&
         Height          =   465
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   30
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "creado por David Guillen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C500&
      Height          =   240
      Left            =   3780
      TabIndex        =   3
      Top             =   4230
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   1710
      Picture         =   "Cargador.frx":0000
      Top             =   2025
      Width           =   3000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00009600&
      Height          =   225
      Left            =   135
      Shape           =   4  'Rounded Rectangle
      Top             =   4545
      Width           =   5820
   End
   Begin VB.Label info 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C500&
      Height          =   240
      Left            =   225
      TabIndex        =   2
      Top             =   4230
      Width           =   2535
   End
End
Attribute VB_Name = "Cargador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub IncrementarBarra()
Static PasoA As Integer
PasoA = PasoA + 1
Barra.Width = (PasoA * BarraMax.Width) / PasosBarra
End Sub

Private Sub Form_Load()
Randomize Timer
End Sub

