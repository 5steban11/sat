VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5490
   ForeColor       =   &H0080FF80&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6225
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   2760
      TabIndex        =   8
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   615
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante\Desktop\Rudy Rodriguez\Ventas De Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pelicula"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "Cod_actor"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Cod_Tipo "
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Cod_Actor "
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Cod_Tipo "
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Pelicula "
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Data1.Recordset.AddNew
End Sub

Private Sub Command3_Click()
Data1.Recordset.Update
End Sub

Private Sub Command4_Click()
Data1.Recordset.Delete
End Sub
