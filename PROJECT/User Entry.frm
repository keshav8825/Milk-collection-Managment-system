VERSION 5.00
Begin VB.Form UserEntry 
   BackColor       =   &H00FF8080&
   Caption         =   "User Entry"
   ClientHeight    =   7485
   ClientLeft      =   5895
   ClientTop       =   2070
   ClientWidth     =   9705
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   9705
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Submit 
      Caption         =   "Submit"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User Id"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mob. No."
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   3960
      Picture         =   "User Entry.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2520
   End
End
Attribute VB_Name = "USERENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Form_Load()
conn
End Sub

Private Sub Submit_Click()

sql = "insert into login values ('" + Text3.Text + "' , '" + Text1.Text + "' ,'" + Text5.Text + "' ,'" + Text2.Text + "')"
If Text4.Text = Text5.Text Then
Set r = c.Execute(sql)
MsgBox "User Created"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Unload Me
LoginForm.Show
Else
MsgBox "new password & confirm password unmatched"
Text4.Text = ""
Text5.Text = ""
Text4.SetFocus
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.Text = UCase(Text1.Text)
Text2.SetFocus
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub
Private Sub text3_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.Text = UCase(Text3.Text)
Text4.SetFocus
End If
End Sub
Private Sub text4_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.Text = UCase(Text4.Text)
Text5.SetFocus
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.Text = UCase(Text5.Text)
Submit.SetFocus '
End If
End Sub
