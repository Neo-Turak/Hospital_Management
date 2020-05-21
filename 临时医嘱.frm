VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form 临时医嘱 
   BackColor       =   &H00FFC0C0&
   Caption         =   "临时医嘱"
   ClientHeight    =   9360
   ClientLeft      =   -30
   ClientTop       =   450
   ClientWidth     =   12480
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
   ScaleHeight     =   16.51
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   22.013
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      Caption         =   "重  置"
      Height          =   495
      Left            =   1800
      TabIndex        =   63
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "查  询"
      Height          =   495
      Left            =   120
      TabIndex        =   62
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Frame Frame6 
      Caption         =   "医嘱状态"
      Height          =   1455
      Left            =   120
      TabIndex        =   59
      Top             =   5880
      Width           =   3255
      Begin VB.OptionButton Option6 
         Caption         =   "未执行"
         Height          =   375
         Left            =   480
         TabIndex        =   61
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         Caption         =   "已执行"
         Height          =   375
         Left            =   480
         TabIndex        =   60
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "病人医嘱"
      Height          =   1455
      Left            =   120
      TabIndex        =   56
      Top             =   4440
      Width           =   3255
      Begin VB.OptionButton Option4 
         Caption         =   "全部医嘱"
         Height          =   495
         Left            =   600
         TabIndex        =   58
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         Caption         =   "这次医嘱"
         Height          =   495
         Left            =   600
         TabIndex        =   57
         Top             =   360
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   240
      Top             =   8040
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "检查项目"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1455
      Left            =   3600
      TabIndex        =   42
      Top             =   2760
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   2566
      _Version        =   393216
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "项目名称"
         Caption         =   "项目名称"
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
      BeginProperty Column02 
         DataField       =   "所属科室"
         Caption         =   "所属科室"
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
         DataField       =   "价格"
         Caption         =   "价格"
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
            ColumnWidth     =   4.313
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1.164
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2.328
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2.99
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2.117
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1455
      Left            =   3600
      TabIndex        =   41
      Top             =   2760
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   2566
      _Version        =   393216
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
      ColumnCount     =   6
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "用法"
         Caption         =   "用法"
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
            ColumnWidth     =   3.466
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4.789
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1.349
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1.296
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1.296
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1.482
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   7320
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "药品库存"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "医嘱专区"
      Height          =   9015
      Left            =   3480
      TabIndex        =   21
      Top             =   240
      Width           =   8775
      Begin VB.CommandButton Command7 
         Caption         =   "加空行"
         Height          =   495
         Left            =   4920
         TabIndex        =   54
         Top             =   2000
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "保存"
         Height          =   495
         Left            =   3720
         TabIndex        =   53
         Top             =   2000
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "快速打印"
         Height          =   855
         Left            =   7920
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   8160
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   45
         Top             =   8160
         Width           =   7695
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   6120
            TabIndex        =   65
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   5280
            TabIndex        =   64
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            Caption         =   "打印"
            Height          =   495
            Left            =   6600
            TabIndex        =   55
            Top             =   200
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   4440
            TabIndex        =   50
            Text            =   "4"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   3480
            TabIndex        =   49
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton Option2 
            Caption         =   "续页打印"
            Height          =   400
            Left            =   1680
            TabIndex        =   47
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "新页打印"
            Height          =   400
            Left            =   240
            TabIndex        =   46
            Top             =   200
            Width           =   1575
         End
         Begin VB.Line Line1 
            X1              =   5040
            X2              =   5040
            Y1              =   120
            Y2              =   720
         End
         Begin VB.Label Label16 
            Caption         =   "从      到"
            Height          =   405
            Left            =   3120
            TabIndex        =   48
            Top             =   250
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "审核"
         Height          =   495
         Left            =   2520
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2000
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "删除"
         Height          =   495
         Left            =   1320
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2000
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "添加"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2000
         Width           =   975
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   3360
         Top             =   7440
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
         RecordSource    =   "临时医嘱"
         Caption         =   "Adodc2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "临时医嘱.frx":0000
         Height          =   3975
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4080
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7011
         _Version        =   393216
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
         ColumnCount     =   16
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
            DataField       =   "类别"
            Caption         =   "类别"
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
            DataField       =   "所配科室"
            Caption         =   "所配科室"
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
         BeginProperty Column04 
            DataField       =   "名称"
            Caption         =   "名称"
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
            DataField       =   "用法"
            Caption         =   "用法"
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
         BeginProperty Column09 
            DataField       =   "金额"
            Caption         =   "金额"
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
            DataField       =   "科室"
            Caption         =   "科室"
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
            DataField       =   "医生"
            Caption         =   "医生"
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
            DataField       =   "医嘱日期"
            Caption         =   "医嘱日期"
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
         BeginProperty Column13 
            DataField       =   "医嘱时间"
            Caption         =   "医嘱时间"
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
         BeginProperty Column14 
            DataField       =   "执行时间"
            Caption         =   "执行时间"
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
         BeginProperty Column15 
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
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   645.165
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text9 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         Height          =   375
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text8 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Y""123"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1600
         Width           =   735
      End
      Begin VB.ComboBox Combo5 
         Height          =   360
         ItemData        =   "临时医嘱.frx":0015
         Left            =   2400
         List            =   "临时医嘱.frx":0022
         TabIndex        =   2
         Text            =   "标志"
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         ItemData        =   "临时医嘱.frx":0032
         Left            =   3360
         List            =   "临时医嘱.frx":0042
         TabIndex        =   3
         Text            =   "用法"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Text            =   "1"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1150
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   840
         TabIndex        =   0
         Top             =   1150
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "临时医嘱.frx":0066
         Left            =   240
         List            =   "临时医嘱.frx":0070
         TabIndex        =   22
         Text            =   "分类"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000010&
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   8415
         Begin MSAdodcLib.Adodc Adodc4 
            Height          =   450
            Left            =   3720
            Top             =   240
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   794
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
            RecordSource    =   "检查科室"
            Caption         =   "Adodc4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "临时医嘱.frx":0084
            Height          =   390
            Left            =   2040
            TabIndex        =   51
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   688
            _Version        =   393216
            ListField       =   "科室名称"
            Text            =   "检查科室"
         End
         Begin VB.CommandButton Command5 
            Caption         =   "导入模板"
            Height          =   495
            Left            =   6600
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   200
            Width           =   1215
         End
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "金额"
         Height          =   255
         Index           =   12
         Left            =   7320
         TabIndex        =   34
         Top             =   2050
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "单价"
         Height          =   255
         Index           =   11
         Left            =   7320
         TabIndex        =   32
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "数  量"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   31
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7800
         TabIndex        =   30
         Top             =   1155
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "库存"
         Height          =   255
         Index           =   6
         Left            =   7320
         TabIndex        =   29
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "规格"
         Height          =   375
         Index           =   3
         Left            =   5280
         TabIndex        =   28
         Top             =   1150
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "规格"
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   27
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   26
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "助记码"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "病人信息"
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   3255
      Begin VB.Timer Timer1 
         Left            =   2400
         Top             =   1200
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1320
         TabIndex        =   40
         Top             =   3600
         Width           =   1680
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "患者编号："
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "入院日期："
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "诊断"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "住院号："
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "床号："
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
   End
   Begin MSForms.Image Image1 
      Height          =   255
      Left            =   1440
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      Size            =   "1931;450"
      Picture         =   "临时医嘱.frx":0099
   End
   Begin VB.Label Label13 
      Caption         =   "科室"
      Height          =   375
      Left            =   600
      TabIndex        =   38
      Top             =   9000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "医生"
      Height          =   375
      Left            =   600
      TabIndex        =   37
      Top             =   8520
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "临时医嘱"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 And Combo2.Text = "药品" Then
 Text2.SetFocus
End If

If KeyAscii = 13 And Combo2.Text = "检查项目" Then
DataCombo1.Visible = True
DataCombo1.SetFocus
Else
End If
End Sub

Private Sub Combo2_LostFocus()
On Error Resume Next
If Combo2.Text = "药品" Then
DataGrid1.Visible = True
DataGrid3.Visible = False
Label9(8).Visible = True
Text5.Visible = True
Combo5.Visible = True
DataCombo1.Visible = False
Combo3.Visible = True

Label9(6).Visible = True
Label10.Visible = True
Label9(12).Visible = True
Text9.Visible = True
Text2.SetFocus
End If

If Combo2.Text = "检查项目" Then
DataGrid3.Visible = True
DataGrid1.Visible = False
Label9(8).Visible = False
Text5.Visible = False
DataCombo1.Visible = True
Label9(6).Visible = False
Label10.Visible = False
Combo5.Visible = False
Combo3.Visible = False

DataCombo1.SetFocus
Label9(12).Visible = False
Text9.Visible = False
End If
End Sub



Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo3.SetFocus
End If
End Sub
Private Sub Command1_Click()
Adodc2.Recordset.Delete
End Sub

Private Sub Command10_Click()
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False

End Sub

Private Sub Command2_Click()
Timer1.Interval = 100
Adodc2.Recordset.MoveFirst
End Sub

Private Sub Command3_Click()
On Error Resume Next
Adodc2.Recordset.MoveFirst
Dim KS As Integer
Dim JS As Integer
Dim XH As Integer
Dim KW As Integer
Dim JW As Integer
KS = Val(Text1.Text)
JS = Val(Text4.Text)
KW = Val(Text6.Text)
JW = Val(Text7.Text)

For XH = 2 To KS Step 1
 Adodc2.Recordset.MoveNext
 Next XH
 
 If Option1.Value = True Then
 
 Printer.PaperSize = 13   'vbPRPSB5 13 B5, 182 x 257 mm
Printer.ScaleMode = vbMillimeters
'标准宽度
Printer.PaintPicture Image1.Picture, 0, 0, 180, 256
Printer.FontSize = 12

Printer.CurrentX = 25
Printer.CurrentY = 36
Printer.Print Left(Label2(0).Caption, 6)

Printer.CurrentX = 63
Printer.CurrentY = 36
Printer.Print Label3.Caption

Printer.CurrentX = 80
Printer.CurrentY = 36
Printer.Print Label4.Caption

Printer.CurrentX = 110
Printer.CurrentY = 36
Printer.Print Label13.Caption

Printer.CurrentX = 140
Printer.CurrentY = 36
Printer.Print Label6.Caption

Printer.CurrentX = 168
Printer.CurrentY = 36
Printer.Print Label5.Caption

'Dim CS As Integer

Do
For i = (KW * 10) + 50 To (JW * 10) + 50 Step 10
Printer.CurrentX = 10
Printer.CurrentY = i
Printer.Print Right(Adodc2.Recordset.Fields("医嘱日期"), 5)
Printer.CurrentX = 23
Printer.CurrentY = i
Printer.Print Left(Adodc2.Recordset.Fields("医嘱时间"), 5)
Printer.CurrentX = 40
Printer.CurrentY = i
If Adodc2.Recordset.Fields("类别") = "检查项目" Then
Printer.Print Adodc2.Recordset.Fields("标志") & "# " & Adodc2.Recordset.Fields("名称") & "/" & Adodc2.Recordset.Fields("规格") & " /" & Adodc2.Recordset.Fields("数量")
Else

If Adodc2.Recordset.Fields("标志") = "┗" Then

Printer.Print Adodc2.Recordset.Fields("标志") & Adodc2.Recordset.Fields("名称") & "/" & Adodc2.Recordset.Fields("规格") & " /" & Adodc2.Recordset.Fields("数量") & Adodc2.Recordset.Fields("用法")
Else
Printer.Print Adodc2.Recordset.Fields("标志") & Adodc2.Recordset.Fields("名称") & "/" & Adodc2.Recordset.Fields("规格") & " /" & Adodc2.Recordset.Fields("数量")
End If

End If

Printer.CurrentX = 135
Printer.CurrentY = i
Printer.Print Left(Adodc2.Recordset.Fields("执行时间"), 5)
Adodc2.Recordset.MoveNext
Printer.DrawStyle = 0
Printer.Line (10, i + 5)-(173, i + 5)
If Adodc2.Recordset.EOF = True Then Exit Do
Next i

Loop
Printer.EndDoc
 End If
 
 If Option2.Value = True Then
 
 Printer.PaperSize = 13   'vbPRPSB5 13 B5, 182 x 257 mm
Printer.ScaleMode = vbMillimeters
'标准宽度
Printer.FontSize = 12

Do
For i = (KW * 10) + 50 To (JW * 10) + 50 Step 10
Printer.CurrentX = 10
Printer.CurrentY = i
Printer.Print Right(Adodc2.Recordset.Fields("医嘱日期"), 5)
Printer.CurrentX = 23
Printer.CurrentY = i
Printer.Print Left(Adodc2.Recordset.Fields("医嘱时间"), 5)
Printer.CurrentX = 40
Printer.CurrentY = i
If Adodc2.Recordset.Fields("类别") = "检查项目" Then
Printer.Print Adodc2.Recordset.Fields("标志") & "# " & Adodc2.Recordset.Fields("名称") & "/" & Adodc2.Recordset.Fields("规格") & " /" & Adodc2.Recordset.Fields("数量")
Else

If Adodc2.Recordset.Fields("标志") = "┗" Then

Printer.Print Adodc2.Recordset.Fields("标志") & Adodc2.Recordset.Fields("名称") & "/" & Adodc2.Recordset.Fields("规格") & " /" & Adodc2.Recordset.Fields("数量") & Adodc2.Recordset.Fields("用法")
Else
Printer.Print Adodc2.Recordset.Fields("标志") & Adodc2.Recordset.Fields("名称") & "/" & Adodc2.Recordset.Fields("规格") & " /" & Adodc2.Recordset.Fields("数量")
End If

End If

Printer.CurrentX = 135
Printer.CurrentY = i
Printer.Print Left(Adodc2.Recordset.Fields("执行时间"), 5)
Adodc2.Recordset.MoveNext
Printer.DrawStyle = 0
Printer.Line (10, i + 5)-(173, i + 5)
If Adodc2.Recordset.EOF = True Then Exit Do
Next i

Loop
Printer.EndDoc
End If
 
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Val(Adodc2.Recordset.RecordCount) > 19 Then
MsgBox "医嘱最多数量为19，请选择分开打印！"
Exit Sub
End If

Printer.PaperSize = 13   'vbPRPSB5 13 B5, 182 x 257 mm
Printer.ScaleMode = vbMillimeters
'标准宽度
Printer.PaintPicture Image1.Picture, 0, 0, 180, 256
Printer.FontSize = 12

'For i = 1 To 18
'Printer.CurrentX = i * 10
'Printer.CurrentY = 1
'Printer.Print i
'Printer.CurrentX = 1
'Printer.CurrentY = i * 10
'Printer.Print i
'Next i
Printer.CurrentX = 25
Printer.CurrentY = 36
Printer.Print Left(Label2(0).Caption, 6)

Printer.CurrentX = 63
Printer.CurrentY = 36
Printer.Print Label3.Caption

Printer.CurrentX = 80
Printer.CurrentY = 36
Printer.Print Label4.Caption

Printer.CurrentX = 110
Printer.CurrentY = 36
Printer.Print Label13.Caption

Printer.CurrentX = 140
Printer.CurrentY = 36
Printer.Print Label6.Caption

Printer.CurrentX = 168
Printer.CurrentY = 36
Printer.Print Label5.Caption

'Dim CS As Integer
Adodc2.Recordset.MoveFirst

Do
For i = 60 To 250 Step 10
Printer.CurrentX = 10
Printer.CurrentY = i
Printer.Print Right(Adodc2.Recordset.Fields("医嘱日期"), 5)
Printer.CurrentX = 23
Printer.CurrentY = i
Printer.Print Left(Adodc2.Recordset.Fields("医嘱时间"), 5)
Printer.CurrentX = 40
Printer.CurrentY = i
If Adodc2.Recordset.Fields("类别") = "检查项目" Then
Printer.Print Adodc2.Recordset.Fields("标志") & "# " & Adodc2.Recordset.Fields("名称") & "/" & Adodc2.Recordset.Fields("规格") & " /" & Adodc2.Recordset.Fields("数量")
Else

If Adodc2.Recordset.Fields("标志") = "┗" Then

Printer.Print Adodc2.Recordset.Fields("标志") & Adodc2.Recordset.Fields("名称") & "/" & Adodc2.Recordset.Fields("规格") & " /" & Adodc2.Recordset.Fields("数量") & Adodc2.Recordset.Fields("用法")
Else
Printer.Print Adodc2.Recordset.Fields("标志") & Adodc2.Recordset.Fields("名称") & "/" & Adodc2.Recordset.Fields("规格") & " /" & Adodc2.Recordset.Fields("数量")
End If

End If

Printer.CurrentX = 135
Printer.CurrentY = i
Printer.Print Left(Adodc2.Recordset.Fields("执行时间"), 5)
Adodc2.Recordset.MoveNext
Printer.DrawStyle = 0
Printer.Line (10, i + 5)-(173, i + 5)
If Adodc2.Recordset.EOF = True Then Exit Do
Next i
Loop
Printer.EndDoc
End Sub

Private Sub Command5_Click()
MsgBox "待开发功能！"
End Sub

Private Sub Command6_Click()
On Error Resume Next
Adodc2.Recordset.UpdateBatch adAffectCurrent
End Sub

Private Sub Command7_Click()
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(3) = ""
Adodc2.Recordset.Fields(4) = ""

End Sub

Private Sub Command8_Click()
On Error Resume Next
With Adodc2.Recordset
.AddNew
.Fields("序号") = Val(.RecordCount) + 1
.Fields("住院号") = Label6.Caption
.Fields("床号") = Label5.Caption
.Fields("姓名") = Label2(0).Caption
.Fields("性别") = Label3.Caption
.Fields("年龄") = Label4.Caption
.Fields("患者编号") = Label15.Caption
.Fields("类别") = Combo2.Text

 If Combo2.Text = "药品" Then
    .Fields("所配科室") = "药房"
    If Combo5.Text = "标志" Then
             Else
    .Fields("标志") = Combo5.Text
            End If
.Fields("名称") = Text3.Text
.Fields("规格") = Label9(3).Caption

    If Combo2.Text = "药品" Then
    .Fields("数量") = Text5.Text
     End If

.Fields("用法") = Combo3.Text
.Fields("单价") = Text8.Text
.Fields("金额") = Text9.Text
End If
If Combo2.Text = "检查项目" Then
.Fields("所配科室") = DataCombo1.Text
.Fields("名称") = Text3.Text
.Fields("数量") = "1"
.Fields("规格") = Label9(3).Caption
.Fields("单价") = Text8.Text
.Fields("金额") = Text8.Text
End If

.Fields("科室") = Label13.Caption
.Fields("医生") = Label12.Caption
If Combo5.Text = "┃" Or Combo5.Text = "┗" Then
.Fields("医嘱日期") = ""
.Fields("医嘱时间") = ""
Else
.Fields("医嘱日期") = Format(Date, "YYYY-MM-DD")
.Fields("医嘱时间") = Format(Time, "HH:MM:SS")
.Fields("状态") = "待审核"
End If
End With

End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Dim SQL As String
SQL = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"

If (Option3.Value = True = True And Option5.Value = True) Then  '这次医嘱，已执行
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
SQL = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from 临时医嘱 where 姓名='" & Label2(0).Caption & "' and 住院号='" & Label6.Caption & "'and 状态='完成' order by 医嘱日期,医嘱时间", Con, adOpenKeyset, adLockOptimistic
Set DataGrid2.DataSource = Mrc
Set Adodc2.Recordset = Mrc
End If

If (Option3.Value = True And Option6.Value = True) Then   '这次医嘱，未执行
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
SQL = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from 临时医嘱 where 姓名='" & Label2(0).Caption & "' and 住院号='" & Label6.Caption & "' and 状态='待执行' order by 医嘱日期,医嘱时间", Con, adOpenKeyset, adLockOptimistic
Set DataGrid2.DataSource = Mrc
Set Adodc2.Recordset = Mrc
End If

If (Option4.Value = True And Option5.Value = True) Then  '全部医嘱，已执行
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from 临时医嘱 where 姓名='" & Label2(0).Caption & "' and 患者编号='" & Label15.Caption & "'and 状态='完成'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid2.DataSource = Mrc
Set Adodc2.Recordset = Mrc
End If
 
If (Option4.Value = True And Option6.Value = True) Then   '全部医嘱，未执行
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from 临时医嘱 where 姓名='" & Label2(0).Caption & "' and 患者编号='" & Label15.Caption & "'and 状态='待执行'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid2.DataSource = Mrc
Set Adodc2.Recordset = Mrc
End If
End Sub

Private Sub DataCombo1_Change()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from 检查项目 where 所属科室 like'%" & DataCombo1.Text & "%'", Con, adOpenKeyset, adLockOptimistic
Set Adodc3.Recordset = Mrc
Set DataGrid3.DataSource = Mrc
DataGrid3.Refresh

End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Form_Activate()
Combo2.SetFocus
End Sub

Private Sub Form_Load()
DataGrid3.Visible = False
DataGrid1.Visible = False
Text4.Text = Adodc2.Recordset.RecordCount
Label12.Caption = 医生工作站MDI.StatusBar1.Panels(3).Text
Label13.Caption = 医生工作站MDI.StatusBar1.Panels(4).Text

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Text4.Text = Adodc2.Recordset.RecordCount
End If
End Sub

Private Sub Text2_Change()
On Error Resume Next
If Combo2.Text = "分类" Then
MsgBox "请选择医嘱分类", vbInformation, "警告"
Combo2.SetFocus
End If
If Combo2.Text = "药品" Then
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from 药品库存 where 助记码 like'%" & Text2.Text & "%'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid1.DataSource = Mrc
Set Adodc1.Recordset = Mrc
DataGrid1.Refresh

Text3.DataField = "药品名"
Set Text3.DataSource = Mrc
 
Set Label9(3).DataSource = Mrc
Set Label10.DataSource = Mrc
 Label10.DataField = "库存"
    Text8.DataField = ""
Set Text8.DataSource = Mrc
 Text8.DataField = "单价"
End If

If Combo2.Text = "检查项目" Then
Dim Conn As ADODB.Connection
Dim mrcc As ADODB.Recordset
Set Conn = New ADODB.Connection
Set mrcc = New ADODB.Recordset
Conn.CursorLocation = adUseClient
SQL = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
Conn.Open SQL
mrcc.Open "select * from 检查项目 where 所属科室='" & DataCombo1.Text & "'and 助记码 like'%" & Text2.Text & "%'", Conn, adOpenKeyset, adLockOptimistic
Set DataGrid3.DataSource = mrcc
Set Adodc3.Recordset = mrcc

Text3.DataField = ""
Set Text3.DataSource = mrcc
Text3.DataField = "项目名称"

Label9(3).DataField = "单位"
Set Label9(3).DataSource = mrcc

   Text8.DataField = ""
Set Text8.DataSource = mrcc
   
   Text8.DataField = "价格"
DataGrid3.Refresh
End If
End Sub

Private Sub Text2_GotFocus()
Text2.Text = ""
Text3.Text = ""
Label9(3).Caption = ""
Label10.Caption = ""
Text8.Text = ""
Text5.Text = ""
Text9.Text = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text5.Visible = False Then
Command8.SetFocus
End If

If KeyAscii = 13 And Text5.Visible = True Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo5.SetFocus
End If
End Sub

Private Sub Text5_LostFocus()
If Val(Label10.Caption) < Val(Text5.Text) Then
MsgBox " 库存数量不足，请调整数量或联系药房人员", vbInformation, "警告"
Text5.SetFocus
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
End If
Text9.Text = Val(Text8.Text) * Val(Text5.Text)
End Sub

Private Sub Timer1_Timer()
Adodc2.Recordset.Fields("状态") = "待执行"
Adodc2.Recordset.Fields("执行时间") = Format(Time, "HH:MM:SS")
Adodc2.Recordset.MoveNext
If Adodc2.Recordset.EOF = True Then
 Timer1.Interval = 0
 MsgBox "医嘱已审核！已提交执行！"
 End If
End Sub
