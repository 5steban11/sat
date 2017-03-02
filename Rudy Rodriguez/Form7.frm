VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form7"
   ScaleHeight     =   6915
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   735
      Left            =   4680
      TabIndex        =   6
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cliente"
      Height          =   735
      Left            =   3960
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Alquiler"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Disco"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Actor"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pelicula"
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tipo de Pelicula"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show

End Sub

Private Sub Command2_Click()
Form2.Show

End Sub

Private Sub Command3_Click()
Form3.Show

End Sub

Private Sub Command4_Click()
Form4.Show


End Sub

Private Sub Command5_Click()
Form5.Show

End Sub

Private Sub Command6_Click()
Form6.Show

End Sub
