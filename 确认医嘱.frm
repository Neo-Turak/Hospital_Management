VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form 确认医嘱 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "确认医嘱"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12375
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4440
      Top             =   5280
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.CommandButton Command3 
      Caption         =   "删除全部(&E)"
      Height          =   495
      Left            =   7320
      TabIndex        =   12
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除选行(&W)"
      Height          =   495
      Left            =   4920
      TabIndex        =   11
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认医嘱(&Q)"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   5880
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8070
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "基本信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.Label Label6 
         Caption         =   "4"
         Height          =   375
         Left            =   9600
         TabIndex        =   9
         Top             =   250
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "录入人："
         Height          =   375
         Index           =   2
         Left            =   8760
         TabIndex        =   8
         Top             =   250
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "4"
         Height          =   375
         Left            =   6720
         TabIndex        =   6
         Top             =   255
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "确认人："
         Height          =   375
         Index           =   1
         Left            =   5880
         TabIndex        =   5
         Top             =   250
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "4"
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   250
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "病区名称："
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   3
         Top             =   250
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   250
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "科室名称："
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   250
         Width           =   1335
      End
   End
End
Attribute VB_Name = "确认医嘱"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
