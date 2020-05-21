VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.MDIForm 医生工作站MDI 
   BackColor       =   &H8000000C&
   Caption         =   "住院医生工作站"
   ClientHeight    =   9210
   ClientLeft      =   -180
   ClientTop       =   1125
   ClientWidth     =   12480
   LinkTopic       =   "MDIForm1"
   Picture         =   "住院医生工作站.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8835
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2858
            TextSave        =   "2016-06-10"
            Object.ToolTipText     =   "日期"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "13:04"
            Object.ToolTipText     =   "时间"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "当前用户"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "部门"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "职位"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8334
            Text            =   "荒地镇卫生院"
            TextSave        =   "荒地镇卫生院"
            Object.ToolTipText     =   "医院名称"
         EndProperty
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
   End
   Begin VB.Menu 医理 
      Caption         =   "医嘱处理(&Q)"
      Index           =   1
      Begin VB.Menu 录入 
         Caption         =   "医嘱录入"
         Index           =   1
         Begin VB.Menu 长期医嘱 
            Caption         =   "长期医嘱"
            Shortcut        =   {F2}
         End
         Begin VB.Menu 临时医嘱 
            Caption         =   "临时医嘱"
            Shortcut        =   {F3}
         End
      End
      Begin VB.Menu 确认 
         Caption         =   "确认医嘱"
         Index           =   2
      End
      Begin VB.Menu 停止 
         Caption         =   "停止医嘱"
         Index           =   3
         Shortcut        =   {F4}
      End
      Begin VB.Menu 我的病人 
         Caption         =   "我的病人"
         Index           =   4
         Shortcut        =   {F5}
      End
      Begin VB.Menu line 
         Caption         =   "_______"
         Index           =   5
      End
      Begin VB.Menu 护理 
         Caption         =   "医嘱护理"
         Index           =   6
      End
   End
   Begin VB.Menu 打印 
      Caption         =   "医嘱模板管理(&W)"
      Begin VB.Menu 长期医嘱模板 
         Caption         =   "长期医嘱模板"
         Index           =   1
      End
      Begin VB.Menu 临时医嘱模板 
         Caption         =   "临时医嘱模板"
         Index           =   2
      End
   End
   Begin VB.Menu 护 
      Caption         =   "护理(&E)"
      Index           =   1
      Begin VB.Menu 护管 
         Caption         =   "护理管理"
         Index           =   2
      End
   End
   Begin VB.Menu 信息 
      Caption         =   "病人信息(&T)"
      Index           =   12
      Begin VB.Menu lline 
         Caption         =   "_______"
         Index           =   1
      End
      Begin VB.Menu 查询 
         Caption         =   "以往病历查询"
         Index           =   2
      End
   End
   Begin VB.Menu 设置 
      Caption         =   "系统设置(&S)"
      Begin VB.Menu 常目 
         Caption         =   "常规项目设置"
         Index           =   1
      End
   End
   Begin VB.Menu 个性 
      Caption         =   "个性化设置(&U)"
      Begin VB.Menu 口令 
         Caption         =   "修改口令"
         Index           =   1
      End
      Begin VB.Menu 皮肤 
         Caption         =   "皮肤设置"
         Index           =   2
      End
   End
End
Attribute VB_Name = "医生工作站MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 查询_Click(Index As Integer)
就诊记录查询.Show
End Sub

Private Sub 长期医嘱_Click()
住院医生工作站.临时医嘱.Show
End Sub


Private Sub 长期医嘱模板_Click(Index As Integer)
长期医嘱模板管理.Show
End Sub

Private Sub 口令_Click(Index As Integer)
密码修改.Show
End Sub

Private Sub 临时医嘱_Click()
录入医嘱.Show
End Sub

Private Sub 临时医嘱模板_Click(Index As Integer)
MsgBox "待开发功能"
End Sub

Private Sub 确认_Click(Index As Integer)
确认医嘱.Show
End Sub

Private Sub 我的病人_Click(Index As Integer)
住院医生工作站.我的病人.Show
End Sub
