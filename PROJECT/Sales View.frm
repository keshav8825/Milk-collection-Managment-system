VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SalesView 
   BackColor       =   &H00FF8080&
   Caption         =   "Sales View"
   ClientHeight    =   10245
   ClientLeft      =   4275
   ClientTop       =   1395
   ClientWidth     =   14295
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
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   10245
   ScaleWidth      =   14295
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   540
      Left            =   13080
      TabIndex        =   9
      Top             =   9120
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Sales View.frx":0000
      Height          =   3495
      Left            =   480
      TabIndex        =   8
      Top             =   6360
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   24
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Bill Product"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "BILL_NO"
         Caption         =   "BILL NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "PR_ID"
         Caption         =   "PRODUCT ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "QTY"
         Caption         =   "QUANTITY"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "RATE"
         Caption         =   "RATE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "AMOUNT"
         Caption         =   "AMOUNT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2174.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2174.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2174.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2174.74
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View"
      Height          =   495
      Left            =   12960
      TabIndex        =   7
      Top             =   3480
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   2295
      Left            =   5040
      TabIndex        =   1
      Top             =   600
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   480
         Left            =   1560
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Submit"
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Bill"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Bill No"
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Sales View.frx":0015
      Height          =   3015
      Left            =   480
      TabIndex        =   0
      Top             =   3000
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   5318
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   24
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Bill Details"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "BILL_NO"
         Caption         =   "BILL NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "BILL_DATE"
         Caption         =   "BILL DATE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "ORDER_NO"
         Caption         =   "ORDER NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "CUST_ID"
         Caption         =   "CUSTOMER ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "TOT_AMT"
         Caption         =   "TOTAL AMOUNT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "DISCOUNT"
         Caption         =   "DISCOUNT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "BILL_AMT"
         Caption         =   "BILL  AMOUNT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "PAID"
         Caption         =   "PAID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "DUES"
         Caption         =   "DUES"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1934.929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   8160
      Top             =   4920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=MOON/ADMIN;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=MOON/ADMIN;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from salebill_details where 1>2"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   9600
      Top             =   5040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=MOON/ADMIN;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=MOON/ADMIN;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * FROM SALEBILL_PR"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SALES VIEW"
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
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "SalesView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
Adodc1.RecordSource = "select * from salebill_details " 'where bill_no=" + Text1.Text + ""
Adodc1.Refresh
'Command2.Visible = True
End If
If Option2.Value = True Then
If Text1.Text = "" Then
MsgBox "Enter Bill No"
Text1.SetFocus
Else
Adodc1.RecordSource = "select * from salebill_details where bill_no=" + Text1.Text + ""
Adodc1.Refresh
Command2.Visible = True
End If
End If
End Sub

Private Sub Command2_Click()
DataGrid2.Visible = True
If Option2.Value = True Then
Adodc2.RecordSource = "select * from salebill_pr where bill_no=" + Text1.Text + ""
Adodc2.Refresh
End If
If Option1.Value = True Then
Adodc2.RecordSource = "select * from salebill_pr "  'where bill_no=" + Text1.Text + ""
Adodc2.Refresh
End If
End Sub

Private Sub Command3_Click()
Unload Me
home.Show
End Sub

Private Sub DataGrid1_Click()
DataGrid2.Visible = True
Adodc2.RecordSource = "select * from salebill_pr where bill_no=" + DataGrid1.Text + ""
Adodc2.Refresh
End Sub

Private Sub Form_Load()
DataGrid2.Visible = False
Command2.Visible = False
Text1.Visible = False
End Sub

Private Sub Option1_Click()
Text1.Text = ""
DataGrid2.Visible = False
Text1.Visible = False
Command2.Visible = False
Adodc1.RecordSource = "select * from salebill_details  where 1>2"
Adodc1.Refresh
End Sub

Private Sub Option2_Click()
Text1.Text = ""
DataGrid2.Visible = False
Text1.Visible = True
Command2.Visible = False
Adodc1.RecordSource = "select * from salebill_details  where 1>2"
Adodc1.Refresh
End Sub

