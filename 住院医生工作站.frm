VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.MDIForm ҽ������վMDI 
   BackColor       =   &H8000000C&
   Caption         =   "סԺҽ������վ"
   ClientHeight    =   9210
   ClientLeft      =   -180
   ClientTop       =   1125
   ClientWidth     =   12480
   LinkTopic       =   "MDIForm1"
   Picture         =   "סԺҽ������վ.frx":0000
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
            Object.ToolTipText     =   "����"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "13:04"
            Object.ToolTipText     =   "ʱ��"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "��ǰ�û�"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "����"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "ְλ"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8334
            Text            =   "�ĵ�������Ժ"
            TextSave        =   "�ĵ�������Ժ"
            Object.ToolTipText     =   "ҽԺ����"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu ҽ�� 
      Caption         =   "ҽ������(&Q)"
      Index           =   1
      Begin VB.Menu ¼�� 
         Caption         =   "ҽ��¼��"
         Index           =   1
         Begin VB.Menu ����ҽ�� 
            Caption         =   "����ҽ��"
            Shortcut        =   {F2}
         End
         Begin VB.Menu ��ʱҽ�� 
            Caption         =   "��ʱҽ��"
            Shortcut        =   {F3}
         End
      End
      Begin VB.Menu ȷ�� 
         Caption         =   "ȷ��ҽ��"
         Index           =   2
      End
      Begin VB.Menu ֹͣ 
         Caption         =   "ֹͣҽ��"
         Index           =   3
         Shortcut        =   {F4}
      End
      Begin VB.Menu �ҵĲ��� 
         Caption         =   "�ҵĲ���"
         Index           =   4
         Shortcut        =   {F5}
      End
      Begin VB.Menu line 
         Caption         =   "_______"
         Index           =   5
      End
      Begin VB.Menu ���� 
         Caption         =   "ҽ������"
         Index           =   6
      End
   End
   Begin VB.Menu ��ӡ 
      Caption         =   "ҽ��ģ�����(&W)"
      Begin VB.Menu ����ҽ��ģ�� 
         Caption         =   "����ҽ��ģ��"
         Index           =   1
      End
      Begin VB.Menu ��ʱҽ��ģ�� 
         Caption         =   "��ʱҽ��ģ��"
         Index           =   2
      End
   End
   Begin VB.Menu �� 
      Caption         =   "����(&E)"
      Index           =   1
      Begin VB.Menu ���� 
         Caption         =   "�������"
         Index           =   2
      End
   End
   Begin VB.Menu ��Ϣ 
      Caption         =   "������Ϣ(&T)"
      Index           =   12
      Begin VB.Menu lline 
         Caption         =   "_______"
         Index           =   1
      End
      Begin VB.Menu ��ѯ 
         Caption         =   "����������ѯ"
         Index           =   2
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "ϵͳ����(&S)"
      Begin VB.Menu ��Ŀ 
         Caption         =   "������Ŀ����"
         Index           =   1
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "���Ի�����(&U)"
      Begin VB.Menu ���� 
         Caption         =   "�޸Ŀ���"
         Index           =   1
      End
      Begin VB.Menu Ƥ�� 
         Caption         =   "Ƥ������"
         Index           =   2
      End
   End
End
Attribute VB_Name = "ҽ������վMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ��ѯ_Click(Index As Integer)
�����¼��ѯ.Show
End Sub

Private Sub ����ҽ��_Click()
סԺҽ������վ.��ʱҽ��.Show
End Sub


Private Sub ����ҽ��ģ��_Click(Index As Integer)
����ҽ��ģ�����.Show
End Sub

Private Sub ����_Click(Index As Integer)
�����޸�.Show
End Sub

Private Sub ��ʱҽ��_Click()
¼��ҽ��.Show
End Sub

Private Sub ��ʱҽ��ģ��_Click(Index As Integer)
MsgBox "����������"
End Sub

Private Sub ȷ��_Click(Index As Integer)
ȷ��ҽ��.Show
End Sub

Private Sub �ҵĲ���_Click(Index As Integer)
סԺҽ������վ.�ҵĲ���.Show
End Sub
