VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form6"
   ScaleHeight     =   7290
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   2400
      TabIndex        =   12
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Guardar"
      Height          =   615
      Left            =   2160
      TabIndex        =   11
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Modificar"
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Crear"
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante\Desktop\Rudy Rodriguez\Ventas De Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cliente "
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "Telefono"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "Direccion "
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "Num_Membresia "
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "telefono"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Direccion "
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nombre"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Num_membresia"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form6"
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
