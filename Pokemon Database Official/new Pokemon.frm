VERSION 5.00
Begin VB.Form Adding 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Adding a New Pokemon"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form3"
   ScaleHeight     =   3840
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text8 
      DataField       =   "National Dex Number"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6120
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      DataField       =   "Evolution"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Asmir\Desktop\School\Computer Science\Pokemon.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pokemon"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   1455
      Left            =   6840
      TabIndex        =   10
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      DataField       =   "Hidden Ability"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Pokemon"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ListBox List2 
      DataField       =   "Secondary Type"
      DataSource      =   "Data1"
      Height          =   1035
      ItemData        =   "new Pokemon.frx":0000
      Left            =   3000
      List            =   "new Pokemon.frx":003A
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ListBox List1 
      DataField       =   "Primary Type"
      DataSource      =   "Data1"
      Height          =   1035
      ItemData        =   "new Pokemon.frx":00C2
      Left            =   1560
      List            =   "new Pokemon.frx":00FC
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "Generation"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "Colour"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "Weight"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "Height (m)"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Colour                 (eg. Blue)"
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Evolution       (eg. First)"
      Height          =   495
      Left            =   4560
      TabIndex        =   18
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Hidden Ability        (eg. Jump)"
      Height          =   495
      Left            =   3120
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Generation           (eg. 1)"
      Height          =   495
      Left            =   1560
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Weight (kg)"
      Height          =   495
      Left            =   5880
      TabIndex        =   15
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Height (m)"
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Secondary Type"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Primary Type"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Pokemon Name (eg. Pikachu)"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Adding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.Delete
Form1.Show
Adding.Hide
End Sub

Private Sub Command2_Click()
MsgBox Text1.Text & " Has been Saved to the Pokemon Database"
Form1.Show
Adding.Hide

End Sub

