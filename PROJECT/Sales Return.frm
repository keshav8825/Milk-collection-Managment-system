VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SalesReturn 
   BackColor       =   &H00FF8080&
   Caption         =   "Sales Return"
   ClientHeight    =   9705
   ClientLeft      =   4275
   ClientTop       =   2730
   ClientWidth     =   14385
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   14385
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Product Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   600
      TabIndex        =   18
      Top             =   4080
      Width           =   11055
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   4800
         TabIndex        =   57
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   6720
         TabIndex        =   53
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   3120
         TabIndex        =   51
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1560
         TabIndex        =   50
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   8520
         TabIndex        =   37
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   8520
         TabIndex        =   35
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove"
         Height          =   375
         Left            =   9720
         TabIndex        =   34
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   495
         Left            =   9720
         TabIndex        =   33
         Top             =   1080
         Width           =   735
      End
      Begin VB.ListBox List6 
         Height          =   2220
         Left            =   7440
         TabIndex        =   32
         Top             =   1560
         Width           =   855
      End
      Begin VB.ListBox List5 
         Height          =   2220
         Left            =   6120
         TabIndex        =   31
         Top             =   1560
         Width           =   975
      End
      Begin VB.ListBox List4 
         Height          =   2220
         Left            =   5160
         TabIndex        =   30
         Top             =   1560
         Width           =   735
      End
      Begin VB.ListBox List3 
         Height          =   2220
         Left            =   3960
         TabIndex        =   29
         Top             =   1560
         Width           =   975
      End
      Begin VB.ListBox List2 
         Height          =   2220
         Left            =   1920
         TabIndex        =   28
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   2220
         Left            =   360
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   7440
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   6120
         TabIndex        =   25
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox Combo4 
         Height          =   480
         Left            =   5160
         TabIndex        =   24
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   3960
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1920
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   480
         Left            =   360
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   9720
         TabIndex        =   20
         Top             =   4320
         Width           =   975
      End
      Begin VB.ListBox List7 
         Height          =   2220
         Left            =   8520
         TabIndex        =   19
         Top             =   1560
         Width           =   975
      End
      Begin VB.ListBox List8 
         Height          =   1500
         Left            =   8640
         TabIndex        =   56
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Dues"
         Height          =   375
         Left            =   4680
         TabIndex        =   58
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Payable Amt"
         Height          =   375
         Left            =   6720
         TabIndex        =   54
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Return Amt"
         Height          =   375
         Left            =   3120
         TabIndex        =   52
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount %"
         Height          =   375
         Left            =   1560
         TabIndex        =   49
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   375
         Left            =   2040
         TabIndex        =   48
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Dues"
         Height          =   375
         Left            =   9720
         TabIndex        =   47
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid"
         Height          =   375
         Left            =   8520
         TabIndex        =   46
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amt"
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   375
         Left            =   8520
         TabIndex        =   44
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   375
         Left            =   7440
         TabIndex        =   43
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   375
         Left            =   6120
         TabIndex        =   42
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   375
         Left            =   5160
         TabIndex        =   41
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   375
         Left            =   1920
         TabIndex        =   40
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Id"
         Height          =   375
         Left            =   360
         TabIndex        =   39
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         Height          =   375
         Left            =   3960
         TabIndex        =   38
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   4455
      Left            =   11880
      TabIndex        =   13
      Top             =   4320
      Width           =   2295
      Begin VB.CommandButton Command6 
         Caption         =   "Exit"
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear"
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save"
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "New"
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Customer Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   11055
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1215
         Left            =   7200
         TabIndex        =   12
         Top             =   1680
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   2143
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"Sales Return.frx":0000
      End
      Begin VB.ComboBox Combo2 
         Height          =   480
         Left            =   3720
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8400
         TabIndex        =   8
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   225574913
         CurrentDate     =   44964
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   3720
         TabIndex        =   7
         Top             =   2400
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   480
         Left            =   3720
         TabIndex        =   6
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1800
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Reason For Return"
         Height          =   375
         Left            =   7560
         TabIndex        =   11
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Bill No."
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   375
         Left            =   7560
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   495
         Left            =   840
         TabIndex        =   3
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Id"
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Return No."
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "SALES RETURN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   4680
      TabIndex        =   63
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label26 
      Caption         =   "Label26"
      Height          =   375
      Left            =   9000
      TabIndex        =   62
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Label22"
      Height          =   375
      Left            =   9000
      TabIndex        =   61
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label21 
      Caption         =   "Label21"
      Height          =   375
      Left            =   9000
      TabIndex        =   60
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   9000
      TabIndex        =   59
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label25 
      Caption         =   "Label25"
      Height          =   495
      Left            =   8400
      TabIndex        =   55
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   6960
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   6960
      X2              =   8160
      Y1              =   4920
      Y2              =   5400
   End
End
Attribute VB_Name = "SalesReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer, X As Integer, Y As Single
'Private Sub Combo1_Click()
'Combo2.Clear
'SQL = "select * from cust_entry where cust_id='" + Combo1.Text + "'"
'Set R = C.Execute(SQL)
'Text2.Text = R.Fields(1)
'Text3.Text = R.Fields(4)
'SQL = "select bill_no from salebill_details where cust_id='" + Combo1.Text + "'"
'Set R = C.Execute(SQL)
'Do While Not R.EOF
'Combo2.AddItem R.Fields(0)
'R.MoveNext
'Loop
'RichTextBox1.Enabled = True
'Combo1.Locked = True
'Combo2.Enabled = True
'End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo2.SetFocus
End If
End Sub

Private Sub Combo2_Click()
Combo3.clear
sql = "select pr_id from salebill_pr where bill_no='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo3.AddItem r.Fields(0)
r.MoveNext
Loop
Combo3.Enabled = True
sql = "select cust_id from salebill_details where bill_no='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Combo1.Text = r.Fields(0)
Label25.Caption = r.Fields(0)
sql = "select name from customer_entry where cust_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text2.Text = r.Fields(0)
sql = "select discount from salebill_details where bill_no='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Text12.Text = r.Fields(0)
RichTextBox1.Enabled = True

sql = "select dues from customer_dues where cust_id='" + Label25.Caption + "'"
Set r = c.Execute(sql)
Text14.Text = r.Fields(0)
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
RichTextBox1.SetFocus
End If
End Sub

Private Sub Combo3_Click()
Label4.Caption = Combo3.Text
sql = "select balance from stock where pr_id='" + Label4.Caption + "'"
Set r = c.Execute(sql)
Label21.Caption = r.Fields(0)
sql = "select pr_name from product_entry where pr_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text4.Text = r.Fields(0)
sql = "select weight from product_entry where pr_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text5.Text = r.Fields(0)
sql = "select unit from product_entry where pr_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Combo4.Text = r.Fields(0)
sql = "select rate from salebill_pr where pr_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text7.Text = r.Fields(0)
sql = "select qty from salebill_pr where pr_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Label18.Caption = r.Fields(0)
Text6.Text = r.Fields(0)
sql = "select amount from salebill_pr where pr_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text8.Text = r.Fields(0)
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Combo4.Enabled = True
Text4.Locked = True
Text5.Locked = True
Text7.Locked = True
Text8.Locked = True
Combo4.Locked = True
Command1.Enabled = True
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub


Private Sub Command1_Click()
If Combo3.Text = "" Or Text6.Text = "" Then
MsgBox "Enter All Fields"
Else
Label22.Caption = Val(Text6.Text)
Label26.Caption = Val(Label21.Caption) - Val(Label22.Caption)
List1.AddItem Combo3.Text
List2.AddItem Text4.Text
List3.AddItem Text5.Text
List4.AddItem Combo4.Text
List5.AddItem Text6.Text
List6.AddItem Text7.Text
List7.AddItem Text8.Text
List8.AddItem Val(Label26.Caption)
Combo3.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo4.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Dim tot As Single
For i = 0 To List7.ListCount - 1
tot = tot + Val(List7.List(i))
Next
Text9.Text = tot
Command2.Enabled = True
Text9.Enabled = True
'Text10.Enabled = True
Text11.Enabled = True
Text9.Locked = True
Text11.Locked = True
Text13.Text = Val(Text9.Text) - (Val(Text9.Text) * Val(Text12.Text) / 100)
Text3.Text = Val(Text14.Text) - Val(Text13.Text)
'Text15.Text = 0
'Text16.Text = 0
'Text17.Text = 0
Text10.Text = ""
Text11.Text = ""
'Text10.SetFocus
End If

End Sub

Private Sub Command2_Click()
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
End Sub

Private Sub Command3_Click()
'Combo1.Clear
sql = "select count(return_no) from sales_return"
Set r = c.Execute(sql)
Text1.Text = a & r.Fields(0) + 1
sql = "select bill_no from salebill_details"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo2.AddItem r.Fields(0)
r.MoveNext
Loop
Text1.Enabled = True
Text1.Locked = True
DTPicker1.Enabled = True
Combo2.Enabled = True
Text2.Enabled = True
'Text3.Enabled = True
Text2.Locked = True
'Text3.Locked = True

Command5.Enabled = True
End Sub

Private Sub Command4_Click()
If Text10.Text = "" Then
MsgBox "All Fields Required"
Else
sql = "insert into sales_return values(" + Text1.Text + ",'" + Format(DTPicker1.Value, "dd mmm yyyy") + "','" + Combo1.Text + "'," + Combo2.Text + ",'" + RichTextBox1.Text + "'," + Text13.Text + "," + Text10.Text + "," + Text11.Text + ")"
Set r = c.Execute(sql)
For k = 0 To List1.ListCount - 1
sql = "insert into salesreturn_pr values(" + Text1.Text + ",'" + List1.List(k) + "'," + List5.List(k) + "," + List6.List(k) + "," + List7.List(k) + ")"
Set r = c.Execute(sql)

sql = "update stock set balance =" + List8.List(k) + ""
Set r = c.Execute(sql)
Next
sql = "UPDATE customer_dues SET dues =" + Text11.Text + " WHERE cust_id = '" + Label25.Caption + "'"
Set r = c.Execute(sql)

MsgBox "Record Saved"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
'Text15.Text = ""
'Text16.Text = ""
'Text17.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
RichTextBox1.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear

Text1.Enabled = False
Text2.Enabled = False
'Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
'Text10.Enabled = False
Text11.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
'Command6.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
DTPicker1.Enabled = False
RichTextBox1.Enabled = False
End If
End Sub

Private Sub Command5_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
RichTextBox1.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
'Command6.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
DTPicker1.Enabled = False
RichTextBox1.Enabled = False

End Sub

Private Sub Command6_Click()
Unload Me
home.Show
End Sub

Private Sub Form_Load()
conn
'Combo4.AddItem "Gm"
'Combo4.AddItem "Kg"
'Combo4.AddItem "Ml"
'Combo4.AddItem "Ltr"
'SQL = "select pr_id from product_entry"
'Set R = C.Execute(SQL)
'Do While Not R.EOF
'Combo3.AddItem R.Fields(0)
'R.MoveNext
'Loop
sql = "select cust_id from customer_entry"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop

Text1.Enabled = False
Text2.Enabled = False
'Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
'Text10.Enabled = False
Text11.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
'Command6.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
DTPicker1.Enabled = False
RichTextBox1.Enabled = False
End Sub



Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
RichTextBox1.Text = UCase(RichTextBox1.Text)
Combo3.SetFocus
End If

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter numeric value only")
End If
If KeyAscii = 13 Then
Text11.Text = Val(Text3.Text) - Val(Text10.Text)
Command4.Enabled = True
Command4.SetFocus
End If
End Sub

Private Sub Text10_LostFocus()
Text11.Text = Val(Text10.Text) + Val(Text3.Text)
End Sub



Private Sub Text16_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter numeric value only")
End If
If KeyAscii = 13 Then
'Text17.Text = Val(Text15.Text) - Val(Text16.Text)
Command4.Enabled = True
Command4.SetFocus
End If
End Sub

Private Sub Text16_LostFocus()
Text11.Text = Val(Text15.Text) - Val(Text16.Text)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter numeric value only")
End If

If KeyAscii = 13 Then
Command1.SetFocus
End If
'X = Text6.Text
'Y = Text7.Text
'Text8.Text = X * Y
End Sub

Private Sub Text7_LostFocus()
X = Val(Text6.Text)
Y = Val(Text7.Text)
Text8.Text = X * Y
End Sub
Private Sub Text6_LostFocus()
If Text6.Text > Val(Label18.Caption) Then
MsgBox "Quantity Exceeded"
Text6.SetFocus
Else
X = Val(Text6.Text)
Y = Val(Text7.Text)
Text8.Text = X * Y
End If
End Sub


