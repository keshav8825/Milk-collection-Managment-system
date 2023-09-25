VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form StockView 
   BackColor       =   &H00FF8080&
   Caption         =   "Stock View"
   ClientHeight    =   8475
   ClientLeft      =   3960
   ClientTop       =   2400
   ClientWidth     =   14760
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   14760
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   540
      Left            =   13200
      TabIndex        =   8
      Top             =   6600
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Stock.frx":0000
      Height          =   3495
      Left            =   2040
      TabIndex        =   7
      Top             =   3720
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "PR_ID"
         Caption         =   "PRODUCT ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "PR_NAME"
         Caption         =   "PRODUCT NAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "MFD"
         Caption         =   "MFD"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "EXP"
         Caption         =   "EXP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "BALANCE"
         Caption         =   "BALANCE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2369.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2564.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2324.977
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2250.142
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1814.74
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Submit"
      Height          =   495
      Left            =   6720
      MaskColor       =   &H0000FF00&
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H0000FF00&
      Height          =   2415
      Left            =   4560
      TabIndex        =   1
      Top             =   840
      Width           =   5655
      Begin VB.OptionButton Option2 
         Caption         =   "Product Id"
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1920
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Product"
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "Submit"
         Height          =   495
         Left            =   2160
         MaskColor       =   &H0000FF00&
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8640
      Top             =   4680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
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
      RecordSource    =   "select * from stock WHERE 1=2"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK"
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
      Left            =   6840
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "StockView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
conn
Adodc1.RecordSource = "select * from STOCK"
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
conn
Adodc1.RecordSource = "select * from STOCK where pr_id='" + Combo1.Text + "'"

Adodc1.Refresh
End Sub

Private Sub Command3_Click()
Unload Me
home.Show
End Sub

Private Sub Option1_Click()
Combo1.Visible = False
Command2.Visible = False
Command1.Visible = True
End Sub

Private Sub Option2_Click()
Combo1.clear
Combo1.Visible = True
Command1.Visible = False
Command2.Visible = True
conn
sql = "select pr_id from product_entry"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub
