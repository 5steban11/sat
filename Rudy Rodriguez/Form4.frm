VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4995
   ForeColor       =   &H0080FF80&
   LinkTopic       =   "Form4"
   ScaleHeight     =   7680
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   2760
      TabIndex        =   12
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   2520
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante\Desktop\Rudy Rodriguez\Ventas De Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Disco"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      DataField       =   "Formato"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text3 
      DataField       =   "Cod_Pelicula"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text2 
      DataField       =   "Num_copias"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataField       =   "Codigo"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Formato"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Cod_pelicula"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Num_copias"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Codigo"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Disco"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
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
