VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6285
   ForeColor       =   &H0080FF80&
   LinkTopic       =   "Form3"
   ScaleHeight     =   7860
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   2880
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   2880
      TabIndex        =   9
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   615
      Left            =   720
      TabIndex        =   8
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   615
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante\Desktop\Rudy Rodriguez\Ventas De Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   855
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Actor"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "Fecha_nac"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "Codigo"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha_nac"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Codigo"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Actor"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
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
