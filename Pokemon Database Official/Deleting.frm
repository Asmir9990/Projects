VERSION 5.00
Begin VB.Form Deleting 
   BackColor       =   &H000000FF&
   Caption         =   "Deleting An Entry"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7785
   LinkTopic       =   "Form2"
   ScaleHeight     =   6390
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Last"
      Height          =   495
      Left            =   5400
      TabIndex        =   25
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Next"
      Height          =   495
      Left            =   3960
      TabIndex        =   24
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Previous"
      Height          =   495
      Left            =   2520
      TabIndex        =   23
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "First"
      Height          =   495
      Left            =   1080
      TabIndex        =   22
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      DataField       =   "National Dex Number"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      DataField       =   "Secondary Type"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4560
      TabIndex        =   19
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      DataField       =   "Primary Type"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Asmir\Desktop\School\Computer Science\Pokemon.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pokemon"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "Height (m)"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "Weight"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "Colour"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "Generation"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Pokemon"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "Hidden Ability"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
      Height          =   1455
      Left            =   2880
      TabIndex        =   2
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      DataField       =   "Evolution"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   3840
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
   Begin VB.Label Label10 
      BackColor       =   &H000000FF&
      Caption         =   "National Dex Number"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Pokemon Name"
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Primary Type"
      Height          =   495
      Left            =   3120
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Secondary Type"
      Height          =   495
      Left            =   4560
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Height (m)"
      Height          =   495
      Left            =   6000
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Caption         =   "Weight (kg)"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "Generation"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000FF&
      Caption         =   "Hidden Ability"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H000000FF&
      Caption         =   "Evolution"
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      Caption         =   "Colour"
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "Deleting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nam

Private Sub Command1_Click()
Form1.Show
Deleting.Hide

End Sub

Private Sub Command2_Click()
nam = Data1.Recordset.Fields("pokemon")
Data1.Recordset.Delete
Data1.Refresh
Data1.Recordset.MoveFirst
MsgBox nam & " Has Been Deleted"

End Sub

Private Sub Command3_Click()
Data1.Recordset.MoveFirst

End Sub

Private Sub Command4_Click()
If (Data1.Recordset.BOF <> True) Then
    Data1.Recordset.MovePrevious
End If
End Sub

Private Sub Command5_Click()
If (Data1.Recordset.EOF <> True) Then
    Data1.Recordset.MoveNext
End If
End Sub

Private Sub Command6_Click()
Data1.Recordset.MoveLast
End Sub
