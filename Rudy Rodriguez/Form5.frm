VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form5"
   ScaleHeight     =   7995
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      DataField       =   "Cantidad"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2640
      TabIndex        =   18
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text6 
      DataField       =   "Valor_alquiler "
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2640
      TabIndex        =   17
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text5 
      DataField       =   "Fecha_devolucion "
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2640
      TabIndex        =   16
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   2400
      TabIndex        =   12
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Guardar"
      Height          =   615
      Left            =   2400
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Modificar"
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Crear"
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   5760
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
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Alquiler"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "Fecha_alquiler "
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "Cod_cliente"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "Cod_disco"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "Codigo"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Valor_alquiler"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha_devolucion"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha_alquiler"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cod_cliente"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cod_disco"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
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
      Caption         =   "Alquiler"
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
Attribute VB_Name = "Form5"
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
