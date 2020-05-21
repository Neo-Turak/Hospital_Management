VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form 手写病历 
   Caption         =   "手写病历"
   ClientHeight    =   10695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16665
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10695
   ScaleWidth      =   16665
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      TabIndex        =   33
      Text            =   "Text8"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "病人信息"
      Height          =   735
      Left            =   0
      TabIndex        =   18
      Top             =   120
      Width           =   12375
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10320
         TabIndex        =   30
         Text            =   "Text7"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   28
         Text            =   "Text6"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   26
         Text            =   "5"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Text            =   "3"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   19
         Text            =   "4"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "住院日期："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "住院号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "姓名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "性别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "年龄："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "床位号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   450
      Left            =   10680
      Top             =   9360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "住院病历"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   10680
      Top             =   4200
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
      RecordSource    =   "书写病历模板"
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
   Begin VB.Frame Frame3 
      Caption         =   "病历记录"
      Height          =   4815
      Left            =   8760
      TabIndex        =   16
      Top             =   4560
      Width           =   6495
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "手写病历.frx":0000
         Height          =   4335
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   7646
         _Version        =   393216
         BackColor       =   8388736
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
         ColumnCount     =   12
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
         BeginProperty Column02 
            DataField       =   "主诉"
            Caption         =   "主诉"
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
            DataField       =   "现病史"
            Caption         =   "现病史"
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
            DataField       =   "既往史"
            Caption         =   "既往史"
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
            DataField       =   "体格检查"
            Caption         =   "体格检查"
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
            DataField       =   "辅助检查"
            Caption         =   "辅助检查"
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
            DataField       =   "诊疗意见"
            Caption         =   "诊疗意见"
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
            DataField       =   "医嘱建议"
            Caption         =   "医嘱建议"
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
            DataField       =   "诊疗医师"
            Caption         =   "诊疗医师"
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
            DataField       =   "就诊日期"
            Caption         =   "就诊日期"
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
            DataField       =   "就诊时间"
            Caption         =   "就诊时间"
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
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar4 
      Height          =   690
      Left            =   3000
      TabIndex        =   1
      Top             =   9960
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1217
      ButtonWidth     =   1852
      ButtonHeight    =   1058
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "保存为模板"
            Key             =   "保存为模板"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "保存"
            Key             =   "保存"
            Object.ToolTipText     =   "保存"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "清空"
            Key             =   "列表"
            Object.ToolTipText     =   "列表"
            ImageKey        =   "View List"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "导出WORD"
            Key             =   "打印"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin RichTextLib.RichTextBox RichTextBox4 
      Height          =   1335
      Left            =   840
      TabIndex        =   2
      Top             =   5640
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2355
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"手写病历.frx":0015
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2040
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3598
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"手写病历.frx":009F
   End
   Begin RichTextLib.RichTextBox RichTextBox3 
      Height          =   960
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1693
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"手写病历.frx":0129
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   1095
      Left            =   840
      TabIndex        =   5
      Top             =   4440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1931
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"手写病历.frx":01B3
   End
   Begin RichTextLib.RichTextBox RichTextBox5 
      Height          =   1335
      Left            =   840
      TabIndex        =   6
      Top             =   7080
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2355
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"手写病历.frx":023D
   End
   Begin RichTextLib.RichTextBox RichTextBox6 
      Height          =   1455
      Left            =   840
      TabIndex        =   7
      Top             =   8520
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2566
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"手写病历.frx":02C7
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      MaxLength       =   50
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   4815
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "手写病历.frx":0351
      Height          =   2895
      Left            =   8760
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1200
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   16776960
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
      BeginProperty Column01 
         DataField       =   "主诉"
         Caption         =   "主诉"
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
         DataField       =   "现病史"
         Caption         =   "现病史"
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
         DataField       =   "既往史"
         Caption         =   "既往史"
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
         DataField       =   "体格检查"
         Caption         =   "体格检查"
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
         DataField       =   "辅助检查"
         Caption         =   "辅助检查"
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
         DataField       =   "诊疗意见"
         Caption         =   "诊疗意见"
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
         DataField       =   "医嘱建议"
         Caption         =   "医嘱建议"
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
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   12480
      TabIndex        =   35
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   495
      Left            =   840
      TabIndex        =   34
      Top             =   10080
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "患者编号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   32
      Top             =   280
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "病历模板"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   9000
      TabIndex        =   31
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label16 
      Caption         =   "主诉："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "诊断:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin MSForms.Label Label10 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   975
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "现病史："
      Size            =   "1720;873"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label10 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Width           =   975
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "既往史"
      Size            =   "1720;873"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label10 
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   6240
      Width           =   615
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "体格检查"
      Size            =   "1085;1085"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label10 
      Height          =   615
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   7320
      Width           =   615
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "辅助检查"
      Size            =   "1085;1085"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label10 
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   8280
      Width           =   615
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "诊疗意见"
      Size            =   "1085;1085"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "手写病历"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataGrid3_Click()
RichTextBox3.Text = DataGrid3.Columns("主诉").CellValue(DataGrid3.Bookmark)
RichTextBox1.Text = DataGrid3.Columns("现病史").CellValue(DataGrid3.Bookmark)
RichTextBox2.Text = DataGrid3.Columns("既往史").CellValue(DataGrid3.Bookmark)
RichTextBox4.Text = DataGrid3.Columns("体格检查").CellValue(DataGrid3.Bookmark)
RichTextBox5.Text = DataGrid3.Columns("辅助检查").CellValue(DataGrid3.Bookmark)
RichTextBox6.Text = DataGrid3.Columns("诊疗意见").CellValue(DataGrid3.Bookmark)
End Sub

Private Sub Form_Activate()
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=TOP-PC;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from 住院病历 where 患者编号 ='" & Text8.Text & "' ", Con, adOpenKeyset, adOpenDynamic
Set Adodc2.Recordset = Mrc
Set DataGrid4.DataSource = Mrc
End Sub

Private Sub Form_Load()
Label9.Caption = 医生工作站MDI.StatusBar1.Panels(6).Text
Label2(1).Caption = 医生工作站MDI.StatusBar1.Panels(3).Text
End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As ADODB.Recordset
Set RS = Adodc2.Recordset
    Dim RSs As ADODB.Recordset
    Set RSs = Adodc1.Recordset
    On Error Resume Next
    Select Case Button.Key
        Case "保存为模板"  '应做:添加 '保存为模板' 按钮代码。
          RSs.AddNew
          RSs!患者编号 = Label19.Caption
          RSs!诊断 = Text1.Text
          RSs!主诉 = RichTextBox3.Text
          RSs!现病史 = RichTextBox1.Text
          RSs!既往史 = RichTextBox2.Text
          RSs!体格检查 = RichTextBox4.Text
          RSs!辅助检查 = RichTextBox5.Text
          RSs!诊疗意见 = RichTextBox6.Text
          RSs!医嘱建议 = Combo1.Text
          RSs.Update
        Case "保存" '应做:添加 '保存' 按钮代码。
          RS.AddNew
          RS!患者编号 = Text8.Text
          RS!诊断 = Text1.Text
          RS!主诉 = RichTextBox3.Text
          RS!现病史 = RichTextBox1.Text
          RS!既往史 = RichTextBox2.Text
          RS!体格检查 = RichTextBox4.Text
          RS!辅助检查 = RichTextBox5.Text
          RS!诊疗意见 = RichTextBox6.Text
          RS!医嘱建议 = Combo1.Text
          RS!就诊医师 = Label22.Caption
          RS!就诊日期 = Date
          RS!就诊时间 = Time
           RS.Update
        Case "列表"
        Text1.Text = ""
        RichTextBox3.Text = ""
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox4.Text = ""
        RichTextBox5.Text = ""
        RichTextBox6.Text = ""
        Case "打印"
        
        
        With 病历打印
Dim wordapp As Word.Application
Dim wordobj As Word.Document
Dim i As Integer
Set wordapp = CreateObject("word.application")
wordapp.Visible = False
Set wordobj = wordapp.Documents.Open(App.Path & "\诊断证明书.docx")
For i = 1 To 2
wordobj.Tables(i).Cell(1, 1).Range.InsertBefore Text:="姓名：" & Text2.Text
wordobj.Tables(i).Cell(1, 2).Range.InsertBefore Text:="性别：" & Text3.Text
wordobj.Tables(i).Cell(1, 3).Range.InsertBefore Text:="年龄：" & Text4.Text
wordobj.Tables(i).Cell(2, 1).Range.InsertBefore Text:="民族：维吾尔族"
wordobj.Tables(i).Cell(2, 2).Range.InsertBefore Text:="科室：" & 医生工作站MDI.StatusBar1.Panels(4).Text
wordobj.Tables(i).Cell(2, 3).Range.InsertBefore Text:="住院号：" & Text6.Text
wordobj.Tables(i).Cell(3, 1).Range.InsertBefore Text:="工作单位地址：新莎车县荒地镇" & Left(Text8.Text, 2) & "村"
wordobj.Tables(i).Cell(4, 1).Range.InsertBefore Text:="入院日期：" & Text7.Text
wordobj.Tables(i).Cell(4, 2).Range.InsertBefore Text:="出院日期：" & DateAdd("d", 6, Text7.Text)
wordobj.Tables(i).Cell(5, 1).Range.InsertBefore Text:=Text1.Text
wordobj.Tables(i).Cell(6, 1).Range.InsertBefore Text:="诊断：" & RichTextBox6.Text
wordobj.Tables(i).Cell(7, 1).Range.InsertBefore Text:="科主任："
wordobj.Tables(i).Cell(7, 2).Range.InsertBefore Text:="医师签名："
wordobj.Tables(i).Cell(7, 3).Range.InsertBefore Text:="患者/家属签名："
Next i
wordobj.SaveAs (App.Path & "\诊断证明书" & Text6.Text & ".docx")
wordobj.Close
wordapp.Quit
Set wordobj = Nothing
Set wordapp = Nothing

End With

  End Select
End Sub
