VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ¼��ҽ�� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ҽ��¼��"
   ClientHeight    =   10650
   ClientLeft      =   345
   ClientTop       =   525
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   12150
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      Caption         =   "¼��ҽ��"
      Height          =   8655
      Left            =   120
      TabIndex        =   36
      Top             =   1560
      Width           =   11655
      Begin VB.CommandButton Command7 
         Caption         =   "����"
         Height          =   495
         Left            =   1440
         TabIndex        =   66
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "�ֿ���ӡ"
         Height          =   495
         Left            =   4800
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ɾ ��"
         Height          =   495
         Left            =   120
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   4080
         TabIndex        =   57
         Top             =   2760
         Width           =   2775
         Begin VB.OptionButton Option2 
            Caption         =   "��ҳ��ӡ"
            Height          =   375
            Left            =   1440
            TabIndex        =   64
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��ͷ��ӡ"
            Height          =   375
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox Text21 
            Height          =   375
            Left            =   2160
            TabIndex        =   60
            Text            =   "18"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox Text20 
            Height          =   375
            Left            =   1440
            TabIndex        =   59
            Text            =   "1"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label16 
            Caption         =   "     ��"
            Height          =   375
            Left            =   1320
            TabIndex        =   61
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label15 
            Caption         =   "��ӡλ�ã�"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "����ģ��"
         Height          =   495
         Left            =   2760
         TabIndex        =   55
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Left            =   120
         Top             =   2520
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         ItemData        =   "¼��ҽ��.frx":0000
         Left            =   5280
         List            =   "¼��ҽ��.frx":000D
         TabIndex        =   7
         Text            =   "��־"
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "�� ��"
         Height          =   495
         Left            =   2760
         TabIndex        =   10
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��  ��"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "���ٴ�ӡ"
         Height          =   495
         Left            =   1440
         TabIndex        =   9
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         DataField       =   "ҩƷ��"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   2280
         TabIndex        =   41
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "¼��ҽ��.frx":001D
         Left            =   120
         List            =   "¼��ҽ��.frx":002A
         TabIndex        =   1
         Text            =   "ִ��Ƶ��"
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "¼��ҽ��.frx":003C
         Left            =   3600
         List            =   "¼��ҽ��.frx":004F
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Text            =   "1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   4560
         TabIndex        =   5
         Text            =   "6"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text16 
         DataField       =   "����"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   5760
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text17 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   38
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   6
         Top             =   2280
         Width           =   495
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   4080
         Top             =   3840
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Bindings        =   "¼��ҽ��.frx":0085
         Height          =   3735
         Left            =   6960
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6588
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "������"
            Caption         =   "������"
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
            DataField       =   "ҩƷ��"
            Caption         =   "ҩƷ��"
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
            DataField       =   "���"
            Caption         =   "���"
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
            DataField       =   "��λ"
            Caption         =   "��λ"
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
            DataField       =   "���"
            Caption         =   "���"
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
            DataField       =   "����"
            Caption         =   "����"
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
            DataField       =   "����"
            Caption         =   "����"
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
            DataField       =   "��ע"
            Caption         =   "��ע"
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
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   734.74
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
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   1200
         TabIndex        =   40
         Top             =   2280
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   138543105
         CurrentDate     =   42501
         MaxDate         =   45658
         MinDate         =   42370
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   138543105
         CurrentDate     =   42501
         MaxDate         =   45658
         MinDate         =   42370
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   5400
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
         RecordSource    =   "����ҽ��"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "¼��ҽ��.frx":009A
         Height          =   4095
         Left            =   120
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   4440
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   20
         BeginProperty Column00 
            DataField       =   "���"
            Caption         =   "���"
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
            DataField       =   "ҽ������"
            Caption         =   "ҽ������"
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
            DataField       =   "���"
            Caption         =   "���"
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
            DataField       =   "��־"
            Caption         =   "��־"
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
            DataField       =   "ҽ������"
            Caption         =   "ҽ������"
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
            DataField       =   "���"
            Caption         =   "���"
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
            DataField       =   "ִ��Ƶ��"
            Caption         =   "ִ��Ƶ��"
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
            DataField       =   "����"
            Caption         =   "����"
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
            DataField       =   "һ������"
            Caption         =   "һ������"
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
            DataField       =   "��ҩ��ʽ"
            Caption         =   "��ҩ��ʽ"
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
            DataField       =   "�ܼ�"
            Caption         =   "�ܼ�"
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
            DataField       =   "ҽ������"
            Caption         =   "ҽ������"
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
            DataField       =   "ҽ��ʱ��"
            Caption         =   "ҽ��ʱ��"
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
            DataField       =   "ִ������"
            Caption         =   "ִ������"
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
            DataField       =   "ֹͣ����"
            Caption         =   "ֹͣ����"
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
            DataField       =   "ֹͣʱ��"
            Caption         =   "ֹͣʱ��"
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
         BeginProperty Column16 
            DataField       =   "����"
            Caption         =   "����"
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
         BeginProperty Column17 
            DataField       =   "ҽ��"
            Caption         =   "ҽ��"
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
         BeginProperty Column18 
            DataField       =   "��ʿ"
            Caption         =   "��ʿ"
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
         BeginProperty Column19 
            DataField       =   "״̬"
            Caption         =   "״̬"
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
               ColumnWidth     =   540.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   540.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   524.976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   615.118
            EndProperty
         EndProperty
      End
      Begin MSForms.Image Image1 
         Height          =   615
         Left            =   6240
         Top             =   1800
         Width           =   495
         Size            =   "873;1085"
         Picture         =   "¼��ҽ��.frx":00AF
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�����ƣ�"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   54
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ���"
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   53
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "���"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   3600
         TabIndex        =   52
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "һ��������"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��ʽ��"
         Height          =   375
         Index           =   7
         Left            =   2400
         TabIndex        =   50
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ�䣺"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   49
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ֹͣʱ�䣺"
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   48
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   3400
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ִ��������   ��"
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   46
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "���ۣ�"
         Height          =   375
         Index           =   0
         Left            =   5160
         TabIndex        =   45
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "�ܼۣ�"
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   44
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "��    �ţ�"
         Height          =   375
         Left            =   3480
         TabIndex        =   43
         Top             =   2280
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   11655
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   9240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text9 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   3840
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1080
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   10800
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   9120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   7920
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   6360
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "��Ժ���ڣ�"
         Height          =   375
         Left            =   8040
         TabIndex        =   30
         Top             =   850
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "���߱�ţ�"
         Height          =   375
         Left            =   5760
         TabIndex        =   28
         Top             =   855
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "���㷽ʽ��"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "��  �ϣ�"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "סԺ������"
         Height          =   375
         Index           =   3
         Left            =   9720
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "���ţ�"
         Height          =   375
         Index           =   2
         Left            =   8520
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "���䣺"
         Height          =   375
         Index           =   1
         Left            =   7320
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "�Ա�"
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "������"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "סԺ��:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   375
      Left            =   5520
      TabIndex        =   56
      Top             =   10320
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   375
      Left            =   9000
      TabIndex        =   35
      Top             =   10320
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   375
      Left            =   7560
      TabIndex        =   34
      Top             =   10320
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   3720
      TabIndex        =   33
      Top             =   10320
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "ҽ��������"
      Height          =   375
      Left            =   2400
      TabIndex        =   32
      Top             =   10320
      Width           =   1215
   End
End
Attribute VB_Name = "¼��ҽ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Timer1.Interval = 100
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command2_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
With Adodc1.Recordset
.Fields("���") = Label10.Caption + 1
.Fields("��������") = Text2.Text
.Fields("����") = Text5.Text
.Fields("סԺ��") = Text1.Text
.Fields("ҽ������") = Text9.Text
.Fields("���") = Text12.Text

If Combo3.Text = "��־" Then
.Fields("��־") = ""
Else
.Fields("��־") = Combo3.Text
End If
.Fields("ҽ������") = Text11.Text
.Fields("���") = Label8.Caption
.Fields("ִ��Ƶ��") = Combo1.Text
.Fields("һ������") = Text14.Text
.Fields("��ҩ��ʽ") = Combo2.Text
.Fields("ҽ������") = DTPicker2.Value
.Fields("ҽ��ʱ��") = Time
.Fields("ִ������") = Text15.Text
.Fields("ֹͣ����") = DTPicker3.Value
.Fields("ֹͣʱ��") = Time
.Fields("����") = Label11.Caption
.Fields("ҽ��") = Label13.Caption
.Fields("״̬") = "�����"

End With
Adodc1.Recordset.UpdateBatch adAffectCurrent
Label10.Caption = Adodc1.Recordset.RecordCount
Text19.Text = Adodc1.Recordset.RecordCount
End Sub

Private Sub Command3_Click()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Dim KS As Integer
Dim JS As Integer
Dim XH As Integer
Dim KW As Integer
Dim JW As Integer
KS = Val(Text1.Text)
JS = Val(Text4.Text)
KW = Val(Text6.Text)
JW = Val(Text7.Text)
On Error Resume Next
 Printer.PaperSize = 13   'vbPRPSB5 13 B5, 182 x 257 mm
Printer.ScaleMode = vbMillimeters
'��׼���
Printer.PaintPicture Image1.Picture, 0, 0, 180, 256
Printer.FontSize = 12
Printer.FontBold = True

Printer.CurrentX = 25
Printer.CurrentY = 27
Printer.Print Text2.Text

Printer.CurrentX = 100
Printer.CurrentY = 27
Printer.Print Text5.Text

Printer.CurrentX = 155
Printer.CurrentY = 27
Printer.Print Text1.Text

Printer.FontSize = 10
Printer.FontBold = False

For i = 60 To 230 Step 10
Printer.CurrentX = 5
Printer.CurrentY = i
Printer.Print Right(Format(Adodc1.Recordset.Fields("ҽ������"), "YYYY-MM-DD"), 5)
Printer.CurrentX = 18
Printer.CurrentY = i
Printer.Print Left(Format(Adodc1.Recordset.Fields("ҽ��ʱ��"), "HH:MM:SS"), 5)
Printer.CurrentX = 30
Printer.CurrentY = i
Printer.Print Adodc1.Recordset.Fields("��־") & Adodc1.Recordset.Fields("ҽ������") & "/" & Adodc1.Recordset.Fields("���") & "/" & Adodc1.Recordset.Fields("һ������") & Adodc1.Recordset.Fields("ִ��Ƶ��")
Printer.CurrentX = 120
Printer.CurrentY = i
Printer.Print Right(Format(Adodc1.Recordset.Fields("ֹͣ����"), "YYYY-MM-DD"), 5)
Printer.Line (10, i + 5)-(170, i + 5)
Adodc1.Recordset.MoveNext

If Adodc1.Recordset.EOF = True Then
Exit For
End If
Next i
Printer.EndDoc
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command5_Click()
On Error Resume Next
����ҽ��ģ�����.Show
With ����ҽ��ģ�����
.Text2.Visible = False
.Command2.Visible = False
.Command5.Visible = True
.Command6.Visible = False
.DataGrid3.Visible = False
.Text3.Visible = False
.Text4.Visible = False
.Text5.Visible = False
.Text8.Visible = False
.Combo1.Visible = False
.Combo2.Visible = False
.Combo3.Visible = False
.Text6.Visible = False
.Text7.Visible = False
.Text8.Visible = False
.Command3.Visible = False
.Command4.Visible = False
.Height = 6615
.Width = 12660
End With
End Sub

Private Sub Command6_Click()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Dim KS As Integer
Dim JS As Integer
Dim XH As Integer
Dim KW As Integer
Dim JW As Integer

KW = Val(Text20.Text)
JW = Val(Text21.Text)

If Option1.Value = True Then

 Printer.PaperSize = 13   'vbPRPSB5 13 B5, 182 x 257 mm
Printer.ScaleMode = vbMillimeters
'��׼���
Printer.PaintPicture Image1.Picture, 0, 0, 180, 256
Printer.FontSize = 12
Printer.FontBold = True
Printer.CurrentX = 25
Printer.CurrentY = 27
Printer.Print Text2.Text
Printer.CurrentX = 100
Printer.CurrentY = 27
Printer.Print Text5.Text
Printer.CurrentX = 155
Printer.CurrentY = 27
Printer.Print Text1.Text
Printer.FontSize = 10
Printer.FontBold = False


For i = (KW * 10) + 50 To (JW * 10) + 50 Step 10
Printer.CurrentX = 5
Printer.CurrentY = i
Printer.Print Right(Format(Adodc1.Recordset.Fields("ҽ������"), "YYYY-MM-DD"), 5)
Printer.CurrentX = 18
Printer.CurrentY = i
Printer.Print Left(Format(Adodc1.Recordset.Fields("ҽ��ʱ��"), "HH:MM:SS"), 5)
Printer.CurrentX = 30
Printer.CurrentY = i
Printer.Print Adodc1.Recordset.Fields("��־") & Adodc1.Recordset.Fields("ҽ������") & "/" & Left(Adodc1.Recordset.Fields("���"), 5) & "/" & Adodc1.Recordset.Fields("һ������") & Adodc1.Recordset.Fields("ִ��Ƶ��")
Printer.CurrentX = 120
Printer.CurrentY = i
Printer.Print Right(Format(Adodc1.Recordset.Fields("ֹͣ����"), "YYYY-MM-DD"), 5)
Printer.Line (10, i + 5)-(170, i + 5)
Adodc1.Recordset.MoveNext
    If Adodc1.Recordset.EOF = True Then
    Exit For
    End If

     Next i
Printer.EndDoc
End If


 If Option2.Value = True Then
 
 Printer.PaperSize = 13   'vbPRPSB5 13 B5, 182 x 257 mm
Printer.ScaleMode = vbMillimeters
'��׼���
Printer.FontSize = 12

For i = (KW * 10) + 50 To (JW * 10) + 50 Step 10
Printer.CurrentX = 5
Printer.CurrentY = i
Printer.Print Right(Format(Adodc1.Recordset.Fields("ҽ������"), "YYYY-MM-DD"), 5)
Printer.CurrentX = 18
Printer.CurrentY = i
Printer.Print Left(Format(Adodc1.Recordset.Fields("ҽ��ʱ��"), "HH:MM:SS"), 5)
Printer.CurrentX = 30
Printer.CurrentY = i
Printer.Print Adodc1.Recordset.Fields("��־") & Adodc1.Recordset.Fields("ҽ������") & "/" & Left(Adodc1.Recordset.Fields("���"), 5) & "/" & Adodc1.Recordset.Fields("һ������") & Adodc1.Recordset.Fields("ִ��Ƶ��")
Printer.CurrentX = 120
Printer.CurrentY = i
Printer.Print Right(Format(Adodc1.Recordset.Fields("ֹͣ����"), "YYYY-MM-DD"), 5)
Printer.Line (10, i + 5)-(170, i + 5)
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Exit For
End If
Next i
Printer.EndDoc
End If

End Sub

Private Sub Command7_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Form_Activate()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from ����ҽ�� where סԺ��='" & Text1.Text & "'", Con, adOpenKeyset, adLockOptimistic
Set Adodc1.Recordset = Mrc
Set DataGrid1.DataSource = Mrc
DataGrid1.Refresh
Label10.Caption = Adodc1.Recordset.RecordCount
Label13.Caption = ҽ������վMDI.StatusBar1.Panels(3).Text
Label11.Caption = ҽ������վMDI.StatusBar1.Panels(4).Text
Label12.Caption = ҽ������վMDI.StatusBar1.Panels(5).Text

End Sub

Private Sub Form_Load()
DTPicker2.Value = Format(Date, "YYYY-MM-DD")
DTPicker3.Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Form_Resize()
Me.Width = 12660
Me.Height = 11070
End Sub

Private Sub Text12_GotFocus()
Text12.SelStart = 0
Text12.SelLength = Len(Text12.Text)
End Sub

Private Sub Text13_Change()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from ҩƷ��� where ������ like'%" & Text13.Text & "%'", Con, adOpenKeyset, adLockOptimistic
Set Adodc2.Recordset = Mrc
Set DataGrid2.DataSource = Mrc
DataGrid2.Refresh
End Sub

Private Sub Text13_GotFocus()
Text13.Text = ""
Text16.Text = ""
Combo1.Text = ""

End Sub

Private Sub Text14_GotFocus()
Text14.SelStart = 0
Text14.SelLength = Len(Text14.Text)

End Sub

Private Sub Text14_LostFocus()
'Text17.Text = Val(Combo1.ItemData(Combo1.ListIndex)) * Val(Text14.Text) * Val(Text16.Text)
End Sub

Private Sub Text15_GotFocus()
Text15.SelStart = 0
Text15.SelLength = Len(Text15.Text)
End Sub

Private Sub Text15_LostFocus()
DTPicker3.Value = DateAdd("d", Val(Text15.Text), Date)
End Sub

Private Sub Text21_LostFocus()
If Val(Text21.Text) > 18 Then
MsgBox " �������1��18��ҽ����"
Text21.Text = "18"
End If
End Sub

Private Sub Timer1_Timer()
Adodc1.Recordset.Fields("״̬") = "��ִ��"
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
 Timer1.Interval = 0
 MsgBox "ҽ������ˣ��ύִ�У�"
 End If
End Sub
