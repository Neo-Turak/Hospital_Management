VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form 长期医嘱模板管理 
   Caption         =   "长期医嘱模板管理"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14745
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   14745
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   6720
   End
   Begin VB.TextBox Text8 
      Height          =   405
      Left            =   11040
      TabIndex        =   19
      Text            =   "1"
      Top             =   6600
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   1200
      Top             =   7920
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   661
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   4440
      TabIndex        =   18
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "保存"
      Height          =   615
      Left            =   8160
      TabIndex        =   16
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "导入到当前医嘱"
      Height          =   495
      Left            =   9480
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "删除"
      Height          =   615
      Left            =   6360
      TabIndex        =   14
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "添  加"
      Height          =   645
      Left            =   4560
      TabIndex        =   13
      Top             =   8160
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      Height          =   405
      ItemData        =   "长期医嘱模板管理.frx":0000
      Left            =   8040
      List            =   "长期医嘱模板管理.frx":0013
      TabIndex        =   12
      Text            =   "给药方式"
      Top             =   7560
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   6600
      TabIndex        =   11
      Text            =   "1"
      Top             =   7560
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      Height          =   405
      ItemData        =   "长期医嘱模板管理.frx":0049
      Left            =   4560
      List            =   "长期医嘱模板管理.frx":0059
      TabIndex        =   10
      Text            =   "执行频率"
      Top             =   7560
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   405
      ItemData        =   "长期医嘱模板管理.frx":0081
      Left            =   11760
      List            =   "长期医嘱模板管理.frx":008E
      TabIndex        =   9
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   8760
      TabIndex        =   8
      Top             =   6600
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   6360
      TabIndex        =   7
      Top             =   6600
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   5880
      TabIndex        =   6
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "添加"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7560
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "长期医嘱模板"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   5175
      Left            =   4440
      TabIndex        =   3
      Top             =   840
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9128
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "药品名称"
         Caption         =   "药品名称"
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
         DataField       =   "规格"
         Caption         =   "规格"
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
         DataField       =   "组号"
         Caption         =   "组号"
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
         DataField       =   "标志"
         Caption         =   "标志"
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
         DataField       =   "执行频率"
         Caption         =   "执行频率"
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
         DataField       =   "数量"
         Caption         =   "数量"
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
         DataField       =   "给药方式"
         Caption         =   "给药方式"
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
         DataField       =   "备注"
         Caption         =   "备注"
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
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查  询"
      Height          =   525
      Left            =   7680
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "疾病名称"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   4560
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1080
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   " 长期医嘱疾病模板"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   3975
      Left            =   360
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4800
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "助记码"
         Caption         =   "助记码"
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
         DataField       =   "药品名"
         Caption         =   "药品名"
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
         DataField       =   "规格"
         Caption         =   "规格"
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
         DataField       =   "单位"
         Caption         =   "单位"
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
         DataField       =   "库存"
         Caption         =   "库存"
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
         DataField       =   "批号"
         Caption         =   "批号"
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
         DataField       =   "单价"
         Caption         =   "单价"
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
         DataField       =   "备注"
         Caption         =   "备注"
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
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   780.095
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "长期医嘱模板管理.frx":009E
      Height          =   3375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "疾病名称"
         Caption         =   "疾病名称"
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
            ColumnWidth     =   3314.835
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3.开始添加套餐内容！"
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   10320
      TabIndex        =   24
      Top             =   8400
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2.点击查询！"
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   10320
      TabIndex        =   23
      Top             =   7800
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1.先添加套餐名称！"
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   10320
      TabIndex        =   22
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "【执行频率】  【数量】  【给药方式】    "
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   7200
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "【简码】【ID  【医嘱名称】      【规格】  【组号 【标志】"
      Height          =   375
      Left            =   4560
      TabIndex        =   20
      Top             =   6120
      Width           =   8295
   End
End
Attribute VB_Name = "长期医嘱模板管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=TOP-PC;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from 长期医嘱模板 where 疾病名称 ='" & Text1.Text & "' order by ID", Con, adOpenKeyset, adOpenDynamic
Set Adodc2.Recordset = Mrc
Set DataGrid2.DataSource = Mrc
Text3.Text = Adodc2.Recordset.RecordCount + 1
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("疾病名称") = Text2.Text
DataGrid1.Refresh
End Sub

Private Sub Command3_Click()
On Error Resume Next
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields("ID") = Text3.Text
Adodc2.Recordset.Fields("疾病名称") = Text1.Text
Adodc2.Recordset.Fields("药品名称") = Text4.Text
Adodc2.Recordset.Fields("规格") = Text5.Text
Adodc2.Recordset.Fields("组号") = Text8.Text
Adodc2.Recordset.Fields("标志") = Combo1.Text
Adodc2.Recordset.Fields("执行频率") = Combo2.Text
Adodc2.Recordset.Fields("数量") = Text6.Text
Adodc2.Recordset.Fields("给药方式") = Combo3.Text
Adodc2.Recordset.Update
DataGrid2.Refresh
End Sub

Private Sub Command4_Click()
Adodc2.Recordset.Delete
End Sub

Private Sub Command5_Click()
On Error Resume Next
Adodc2.Recordset.MoveFirst
Dim i As Integer
i = Adodc2.Recordset.RecordCount
For c = 1 To i Step 1
录入医嘱.Adodc1.Recordset.AddNew
录入医嘱.Adodc1.Recordset.Fields("序号") = i + 1
录入医嘱.Adodc1.Recordset.Fields("病人姓名") = 录入医嘱.Text2.Text
录入医嘱.Adodc1.Recordset.Fields("床号") = 录入医嘱.Text5.Text
录入医嘱.Adodc1.Recordset.Fields("住院号") = 录入医嘱.Text1.Text
录入医嘱.Adodc1.Recordset.Fields("医嘱编码") = 录入医嘱.Text9.Text
录入医嘱.Adodc1.Recordset.Fields("组号") = DataGrid2.Columns("组号").CellValue(DataGrid2.Bookmark)
录入医嘱.Adodc1.Recordset.Fields("标志") = DataGrid2.Columns("标志").CellValue(DataGrid2.Bookmark)
录入医嘱.Adodc1.Recordset.Fields("医嘱名称") = DataGrid2.Columns("药品名称").CellValue(DataGrid2.Bookmark)
录入医嘱.Adodc1.Recordset.Fields("规格") = DataGrid2.Columns("规格").CellValue(DataGrid2.Bookmark)
录入医嘱.Adodc1.Recordset.Fields("执行频率") = DataGrid2.Columns("执行频率").CellValue(DataGrid2.Bookmark)
录入医嘱.Adodc1.Recordset.Fields("一次数量") = DataGrid2.Columns("数量").CellValue(DataGrid2.Bookmark)
录入医嘱.Adodc1.Recordset.Fields("给药方式") = DataGrid2.Columns("给药方式").CellValue(DataGrid2.Bookmark)
录入医嘱.Adodc1.Recordset.Fields("医嘱日期") = Format(Date, "YYYY-MM-DD")
录入医嘱.Adodc1.Recordset.Fields("医嘱时间") = Format(Time, "HH:MM:SS")
录入医嘱.Adodc1.Recordset.Fields("执行天数") = "6"
录入医嘱.Adodc1.Recordset.Fields("停止日期") = Format(DateAdd("d", "6", Date), "YYYY-MM-DD")
录入医嘱.Adodc1.Recordset.Fields("科室") = 录入医嘱.Label11.Caption
录入医嘱.Adodc1.Recordset.Fields("医生") = 录入医嘱.Label13.Caption
录入医嘱.Adodc1.Recordset.Fields("状态") = "待审核"
录入医嘱.Adodc1.Recordset.Update
Adodc2.Recordset.MoveNext
录入医嘱.Adodc1.Recordset.UpdateBatch adAffectCurrent
录入医嘱.DataGrid2.Refresh
If Adodc2.Recordset.EOF = True Then
MsgBox "导入长期医嘱模板成功！,", vbInformation, "提示"
Exit Sub
End If
Next c
End Sub

Private Sub Command6_Click()
Adodc2.Recordset.Update
End Sub

Private Sub DataGrid1_DblClick()
Call Command1_Click
End Sub

Private Sub Form_Resize()
Me.Width = 13770
Me.Height = 9840
End Sub

Private Sub Text3_GotFocus()
Text3.Text = Adodc2.Recordset.RecordCount + 1
End Sub

Private Sub Text7_Change()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from 药品库存 where 助记码 like'%" & Text7.Text & "%'", Con, adOpenKeyset, adLockOptimistic
Set Adodc3.Recordset = Mrc
Set DataGrid3.DataSource = Mrc
    DataGrid3.Refresh
Set Text4.DataSource = Mrc
Text4.DataField = "药品名"
Set Text5.DataSource = Mrc
Text5.DataField = "规格"
End Sub

Private Sub Text7_GotFocus()
Text7.Text = ""
Text3.Text = Adodc2.Recordset.RecordCount + 1
Text4.Text = ""
Text5.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
End Sub
