VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Database 
   BackColor       =   &H00FF8080&
   Caption         =   "This Is the Pokemon Database"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11565
   LinkTopic       =   "Form2"
   ScaleHeight     =   10575
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Reset Query"
      Height          =   615
      Left            =   360
      TabIndex        =   30
      Top             =   7800
      Width           =   1575
   End
   Begin VB.ComboBox Combo14 
      Height          =   315
      ItemData        =   "PokeForm2.frx":0000
      Left            =   4800
      List            =   "PokeForm2.frx":000A
      TabIndex        =   8
      Text            =   "And/Or"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.ComboBox Combo13 
      Height          =   315
      ItemData        =   "PokeForm2.frx":0017
      Left            =   1920
      List            =   "PokeForm2.frx":0021
      TabIndex        =   4
      Text            =   "And/Or"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Run Query"
      Height          =   615
      Left            =   3240
      TabIndex        =   28
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FF8080&
      Caption         =   "Sort Three"
      Height          =   1335
      Left            =   6120
      TabIndex        =   27
      Top             =   9000
      Width           =   1575
      Begin VB.ComboBox Combo11 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Text            =   "Field"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox Combo12 
         Height          =   315
         ItemData        =   "PokeForm2.frx":002E
         Left            =   120
         List            =   "PokeForm2.frx":0038
         TabIndex        =   17
         Text            =   "Sort By"
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      Caption         =   "Sort Two"
      Height          =   1335
      Left            =   3240
      TabIndex        =   26
      Top             =   9000
      Width           =   1575
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Text            =   "Field"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         ItemData        =   "PokeForm2.frx":0047
         Left            =   120
         List            =   "PokeForm2.frx":0051
         TabIndex        =   15
         Text            =   "Sort By"
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Sort One"
      Height          =   1335
      Left            =   360
      TabIndex        =   25
      Top             =   9000
      Width           =   1575
      Begin VB.ComboBox Combo8 
         Height          =   315
         ItemData        =   "PokeForm2.frx":0060
         Left            =   120
         List            =   "PokeForm2.frx":006A
         TabIndex        =   13
         Text            =   "Sort By"
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox Combo7 
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "PokeForm2.frx":0079
         Left            =   120
         List            =   "PokeForm2.frx":009B
         TabIndex        =   12
         Text            =   "Field"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Search Three"
      Height          =   1695
      Left            =   6120
      TabIndex        =   24
      Top             =   6120
      Width           =   1455
      Begin VB.ComboBox Combo5 
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "PokeForm2.frx":011F
         Left            =   120
         List            =   "PokeForm2.frx":013E
         TabIndex        =   9
         Text            =   "Field"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "PokeForm2.frx":01AD
         Left            =   120
         List            =   "PokeForm2.frx":01C3
         TabIndex        =   10
         Text            =   "Operation"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Search Two"
      Height          =   1695
      Left            =   3240
      TabIndex        =   23
      Top             =   6000
      Width           =   1455
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "PokeForm2.frx":020A
         Left            =   120
         List            =   "PokeForm2.frx":0229
         TabIndex        =   5
         Text            =   "Field"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "PokeForm2.frx":0298
         Left            =   120
         List            =   "PokeForm2.frx":02AE
         TabIndex        =   6
         Text            =   "Operation"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Search One"
      Height          =   1695
      Left            =   360
      TabIndex        =   22
      Top             =   5880
      Width           =   1455
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "PokeForm2.frx":02F5
         Left            =   120
         List            =   "PokeForm2.frx":030B
         TabIndex        =   2
         Text            =   "Operation"
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "PokeForm2.frx":0352
         Left            =   120
         List            =   "PokeForm2.frx":0371
         TabIndex        =   1
         Text            =   "Field"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Asmir\Desktop\School\Computer Science\Pokemon.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pokemon"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "PokeForm2.frx":03E0
      Height          =   4935
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8705
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   0
      BackColorSel    =   255
      GridColor       =   16744576
      GridColorFixed  =   16744576
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "Select* from Pokemon where "
      Height          =   735
      Left            =   4920
      TabIndex        =   29
      Top             =   7920
      Width           =   6375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Search By"
      Height          =   255
      Left            =   5160
      TabIndex        =   21
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Sort By"
      Height          =   255
      Left            =   3720
      TabIndex        =   20
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   5400
      TabIndex        =   19
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim field1, field2, field3, op1, op2, op3, break1, break2, field4, field5, field6, order1, order2, order3
Private Sub Combo1_Click()
field1 = " [" & Combo1.Text & "] "
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3
End Sub

Private Sub Combo10_Click()
order2 = Combo10.Text
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3 & " Order by " & field4 & order1 & field5 & order2 & field6 & order3

End Sub

Private Sub Combo11_Change()
field6 = Combo11.Text
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3 & " Order by " & field4 & order1 & field5 & order2 & field6 & order3

End Sub

Private Sub Combo12_Change()
order3 = Combo12.Text
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3 & " Order by " & field4 & order1 & field5 & order2 & field6 & order3

End Sub

Private Sub Combo13_click()
break1 = " " & Combo13.Text & " "
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3
End Sub

Private Sub Combo14_Click()
break2 = " " & Combo14.Text & " "
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3
End Sub

Private Sub Combo2_Click()
If Combo2.Text = "Equals" Then
    op1 = " = " & Text1.Text
End If
If Combo2.Text = "Less Then" Then
    op1 = " < " & Text1.Text
End If
If Combo2.Text = "Greater Then" Then
    op1 = " > " & Text1.Text
End If
If Combo2.Text = "Contains" Then
    op1 = " like *" & Text1.Text & "* "
End If
If Combo2.Text = "Starts With" Then
    op1 = " like " & Text1.Text & "* "
End If
If Combo2.Text = "Ends With" Then
    op1 = " like *" & Text1.Text & " "
End If
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3
End Sub

Private Sub Combo3_Click()
field2 = " [" & Combo3.Text & "] "
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3

End Sub

Private Sub Combo4_Click()
If Combo4.Text = "Equals" Then
op2 = " = " & Text2.Text
End If
If Combo4.Text = "Less Then" Then
op2 = " < " & Text2.Text
End If
If Combo4.Text = "Greater Then" Then
op2 = " > " & Text2.Text
End If
If Combo4.Text = "Contains" Then
op2 = " like *" & Text2.Text & "* "
End If
If Combo4.Text = "Starts With" Then
op2 = " like " & Text2.Text & "* "
End If
If Combo4.Text = "Ends With" Then
op2 = " like *" & Text2.Text & " "
End If
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3
End Sub

Private Sub Combo5_Click()
field3 = " [" & Combo5.Text & "] "
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3

End Sub

Private Sub Combo6_Click()
If Combo6.Text = "Equals" Then
op3 = " = " & Text3.Text
End If
If Combo6.Text = "Less Then" Then
op3 = " < " & Text3.Text
End If
If Combo6.Text = "Greater Then" Then
op3 = " > " & Text3.Text
End If
If Combo6.Text = "Contains" Then
op3 = " like *" & Text3.Text & "* "
End If
If Combo6.Text = "Starts With" Then
op3 = " like " & Text3.Text & "* "
End If
If Combo6.Text = "Ends With" Then
op3 = " like *" & Text3.Text & " "
End If
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3

End Sub

Private Sub Combo7_Click()
field4 = Combo7.Text
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3 & " Order by " & field4 & order1 & field5 & order2 & field6 & order3


End Sub

Private Sub Combo8_Click()
order1 = Combo8.Text
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3 & " Order by " & field4 & order1 & field5 & order2 & field6 & order3

End Sub

Private Sub Combo9_Click()
field5 = Combo9.Text
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3 & " Order by " & field4 & order1 & field5 & order2 & field6 & order3

End Sub

Private Sub Command1_Click()
Form1.Show
Database.Hide

End Sub

Private Sub List1_Click()
Data1.Recordset.MoveLast
cnt = Data1.Recordset.Fields("National Dex Number")

sm = Array()

Data1.Recordset.MoveFirst
For abc = 1 To cnt
    For xyz = 0 To cnt
        If (sm(xyz) <> Data1.Recordset.Fields(List1.Text)) Then
            sm(xyz) = Data1.Recordset.Fields(List1.Text)
        End If
    Next xyz
    Data1.Recordset.MoveNext
Next abc

List1.Enabled = False
List2.Visible = True

For abc = 0 To sm.Count
    List2.AddItem (sm(abc))
Next abc
End Sub

Private Sub Command2_Click()
Data1.RecordSource = Label5.Caption
Data1.Refresh

End Sub


Private Sub Command3_Click()
field1 = ""
field2 = ""
field3 = ""
op1 = ""
op2 = ""
op3 = ""
break1 = ""
break2 = ""
field4 = ""
field5 = ""
field6 = ""
order1 = ""
order2 = ""
order3 = ""
Label5.Caption = "Select* from Pokemon where"

End Sub

Private Sub Text1_Change()
If Combo2.Text = "Equals" Then
    op1 = " = " & Text1.Text
End If
If Combo2.Text = "Less Then" Then
    op1 = " < " & Text1.Text
End If
If Combo2.Text = "Greater Then" Then
    op1 = " > " & Text1.Text
End If
If Combo2.Text = "Contains" Then
    op1 = " like *" & Text1.Text & "* "
End If
If Combo2.Text = "Starts With" Then
    op1 = " like " & Text1.Text & "* "
End If
If Combo2.Text = "Ends With" Then
    op1 = " like *" & Text1.Text & " "
End If
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3

End Sub

Private Sub Text2_Change()
If Combo4.Text = "Equals" Then
op2 = " = " & Text2.Text
End If
If Combo4.Text = "Less Then" Then
op2 = " < " & Text2.Text
End If
If Combo4.Text = "Greater Then" Then
op2 = " > " & Text2.Text
End If
If Combo4.Text = "Contains" Then
op2 = " like *" & Text2.Text & "* "
End If
If Combo4.Text = "Starts With" Then
op2 = " like " & Text2.Text & "* "
End If
If Combo4.Text = "Ends With" Then
op2 = " like *" & Text2.Text & " "
End If
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3

End Sub

Private Sub Text3_Change()
If Combo6.Text = "Equals" Then
    op3 = " = " & Text3.Text
End If
If Combo6.Text = "Less Then" Then
    op3 = " < " & Text3.Text
End If
If Combo6.Text = "Greater Then" Then
    op3 = " > " & Text3.Text
End If
If Combo6.Text = "Contains" Then
    op3 = " like *" & Text3.Text & "* "
End If
If Combo6.Text = "Starts With" Then
op3 = " like " & Text3.Text & "* "
End If
If Combo6.Text = "Ends With" Then
op3 = " like *" & Text3.Text & " "
End If
Label5.Caption = "Select* from Pokemon where " & field1 & op1 & break1 & field2 & op2 & break2 & field3 & op3

End Sub
