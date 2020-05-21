VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form 我的病人 
   Caption         =   "我的病人"
   ClientHeight    =   9300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11010
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   11010
   Begin VB.CommandButton Command9 
      Caption         =   "病人入院"
      Height          =   495
      Left            =   9240
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "医嘱管理"
      Height          =   495
      Left            =   9240
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "编写病历"
      Height          =   495
      Left            =   9240
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "病人出院"
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "临时遗嘱"
      Height          =   495
      Left            =   9240
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "长期遗嘱"
      Height          =   495
      Left            =   9240
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "我的病人.frx":0000
      Height          =   8295
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   14631
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   4227327
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "患者编号"
         Caption         =   "患者编号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "床位号"
         Caption         =   "床位号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "患者姓名"
         Caption         =   "患者姓名"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "性别"
         Caption         =   "性别"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "年龄"
         Caption         =   "年龄"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "住院号"
         Caption         =   "住院号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "诊疗医生"
         Caption         =   "诊疗医生"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "诊断"
         Caption         =   "诊断"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "入院日期"
         Caption         =   "入院日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "天数"
         Caption         =   "天数"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "所属医生"
         Caption         =   "所属医生"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "出院日期"
         Caption         =   "出院日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "状态"
         Caption         =   "状态"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2174.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   689.953
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2160
      Top             =   7680
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "床位动态"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "职位"
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   8760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "部门"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   8760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      DataField       =   "用户名"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   8760
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "我的病人"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
录入医嘱.Text2.Text = DataGrid1.Columns("患者姓名").CellValue(DataGrid1.Bookmark)
录入医嘱.Text3.Text = DataGrid1.Columns("性别").CellValue(DataGrid1.Bookmark)
录入医嘱.Text4.Text = DataGrid1.Columns("年龄").CellValue(DataGrid1.Bookmark)
录入医嘱.Text5.Text = DataGrid1.Columns("床位号").CellValue(DataGrid1.Bookmark)

录入医嘱.Text1.Text = DataGrid1.Columns("住院号").CellValue(DataGrid1.Bookmark)
录入医嘱.Text7.Text = DataGrid1.Columns("诊断").CellValue(DataGrid1.Bookmark)
录入医嘱.Text10.Text = DataGrid1.Columns("入院日期").CellValue(DataGrid1.Bookmark)
录入医嘱.Text9.Text = DataGrid1.Columns("患者编号").CellValue(DataGrid1.Bookmark)
Unload Me
录入医嘱.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
临时医嘱.Label2(0).Caption = DataGrid1.Columns("患者姓名").CellValue(DataGrid1.Bookmark)
临时医嘱.Label3.Caption = DataGrid1.Columns("性别").CellValue(DataGrid1.Bookmark)
临时医嘱.Label4.Caption = DataGrid1.Columns("年龄").CellValue(DataGrid1.Bookmark)
临时医嘱.Label5.Caption = DataGrid1.Columns("床位号").CellValue(DataGrid1.Bookmark)

临时医嘱.Label6.Caption = DataGrid1.Columns("住院号").CellValue(DataGrid1.Bookmark)
临时医嘱.Label8.Caption = DataGrid1.Columns("诊断").CellValue(DataGrid1.Bookmark)
临时医嘱.Label7.Caption = DataGrid1.Columns("入院日期").CellValue(DataGrid1.Bookmark)
临时医嘱.Label15.Caption = DataGrid1.Columns("患者编号").CellValue(DataGrid1.Bookmark)
Unload Me
临时医嘱.Show
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
医嘱浏览.Show
End Sub

Private Sub Command6_Click()
On Error Resume Next
手写病历.Text2.Text = DataGrid1.Columns("患者姓名").CellValue(DataGrid1.Bookmark)
手写病历.Text3.Text = DataGrid1.Columns("性别").CellValue(DataGrid1.Bookmark)
手写病历.Text4.Text = DataGrid1.Columns("年龄").CellValue(DataGrid1.Bookmark)
手写病历.Text5.Text = DataGrid1.Columns("床位号").CellValue(DataGrid1.Bookmark)
手写病历.Text6.Text = DataGrid1.Columns("住院号").CellValue(DataGrid1.Bookmark)
手写病历.Text1.Text = DataGrid1.Columns("诊断").CellValue(DataGrid1.Bookmark)
手写病历.Text7.Text = DataGrid1.Columns("入院日期").CellValue(DataGrid1.Bookmark)
手写病历.Text8.Text = DataGrid1.Columns("患者编号").CellValue(DataGrid1.Bookmark)
Unload Me
手写病历.Show
End Sub

Private Sub Command7_Click()
医嘱浏览.Show
End Sub

Private Sub Command8_Click()
医嘱记录查询.Show
End Sub

Private Sub Command9_Click()
病床分配.Show
End Sub

Private Sub Form_Load()

Label1.Caption = 医生工作站MDI.StatusBar1.Panels(3).Text
Label2.Caption = 医生工作站MDI.StatusBar1.Panels(4).Text
Label3.Caption = 医生工作站MDI.StatusBar1.Panels(5).Text

Dim Conn As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Conn = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim ConnectString As String
ConnectString = "Provider=SQLOLEDB.1;password=sa;Persist Security Info=true;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
Conn.Open ConnectString
Conn.CursorLocation = adUseClient
Mrc.Open "select * from 床位动态 where 所属医生='" & Label1.Caption & "'and 病床分区 like '%" & Label2.Caption & "%'", Conn, adOpenKeyset, adLockOptimistic
Set Adodc1.Recordset = Mrc
Set DataGrid1.DataSource = Mrc
End Sub

Private Sub Form_Resize()
Me.Height = 9390
Me.Width = 11415
End Sub
