VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8715
   ForeColor       =   &H0080FF80&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   4680
      TabIndex        =   8
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   615
      Left            =   1800
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante\Desktop\Rudy Rodriguez\Ventas De Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Tipo De Pelicula"
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      DataField       =   "Categoria"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "Tipo"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Categoria "
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Pelicula"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew

End Sub

Private Sub Command3_Click()
Data1.Recordset.Update
End Sub

Private Sub Command4_Click()
Data1.Recordset.Delete
End Sub
