VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ������ӡ 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00004040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������"
   ClientHeight    =   10170
   ClientLeft      =   10365
   ClientTop       =   750
   ClientWidth     =   10815
   FillColor       =   &H00D1815F&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "��鵥��ӡ.frx":0000
   ScaleHeight     =   10170
   ScaleWidth      =   10815
   Begin VB.CommandButton Command1 
      Caption         =   "�� ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   21
      Top             =   0
      Width           =   1215
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar2 
      Height          =   9735
      Left            =   8400
      TabIndex        =   3
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   17171
      _Version        =   393216
      LargeChange     =   100
      Min             =   100
      Max             =   2000
      Orientation     =   1572864
      SmallChange     =   50
      Value           =   280
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   9375
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   8655
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   7920
         Top             =   0
      End
      Begin VB.PictureBox Pbox1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         ClipControls    =   0   'False
         DragIcon        =   "��鵥��ӡ.frx":7CEB8
         ForeColor       =   &H00404040&
         Height          =   8655
         Left            =   600
         ScaleHeight     =   25
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   7215
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "��鵥��ӡ.frx":7D442
      Left            =   1320
      List            =   "��鵥��ӡ.frx":7D45B
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   9840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   661
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   100
      Min             =   100
      Max             =   2000
      Orientation     =   1572865
      SmallChange     =   50
      Value           =   280
   End
   Begin VB.Label Label4 
      Caption         =   "�������11"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "H:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   4
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   8880
      TabIndex        =   20
      ToolTipText     =   "������飺"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "�������label5"
      Height          =   495
      Left            =   8880
      TabIndex        =   19
      ToolTipText     =   "���������"
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "�����10"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   8880
      TabIndex        =   18
      ToolTipText     =   "����飺"
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "����ʷ9"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   8880
      TabIndex        =   17
      ToolTipText     =   "����ʷ��"
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "�ֲ�ʷ8"
      Height          =   495
      Index           =   8
      Left            =   8880
      TabIndex        =   16
      ToolTipText     =   "�ֲ�ʷ��"
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "���7"
      Height          =   495
      Index           =   7
      Left            =   8880
      TabIndex        =   15
      ToolTipText     =   "��ϣ�"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "סԺ����6"
      Height          =   495
      Index           =   6
      Left            =   8880
      TabIndex        =   14
      ToolTipText     =   "סԺ���ڣ�"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "����5"
      Height          =   495
      Index           =   5
      Left            =   8880
      TabIndex        =   13
      ToolTipText     =   "���ţ�"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "סԺ��4"
      Height          =   495
      Index           =   4
      Left            =   8880
      TabIndex        =   12
      ToolTipText     =   "סԺ�ţ�"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "����3"
      Height          =   495
      Index           =   3
      Left            =   8880
      TabIndex        =   11
      ToolTipText     =   "���䣺"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "�Ա�2"
      Height          =   495
      Index           =   2
      Left            =   8880
      TabIndex        =   10
      ToolTipText     =   "�Ա�"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "���߱��1"
      Height          =   495
      Index           =   1
      Left            =   8880
      TabIndex        =   9
      ToolTipText     =   "���߱�ţ�"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "����4��0��"
      Height          =   495
      Index           =   0
      Left            =   8880
      TabIndex        =   8
      ToolTipText     =   "������"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "ҽԺ����13"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "����ҽ��"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   960
      Picture         =   "��鵥��ӡ.frx":7D486
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "������ӡ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Label1.Caption = Left(Combo1.Text, Len(Combo1.Text) - 1) / 100
End Sub

Private Sub Command1_Click()
Printer.Orientation = 1
Printer.PaperSize = 13
Printer.ScaleMode = vbcentimeter
Printer.ScaleWidth = 18
Printer.ScaleHeight = 25
Printer.CurrentX = 6
Printer.CurrentY = 3
Printer.ForeColor = vbRed
Printer.FontBold = True

Printer.FontSize = 16
Printer.Print "�� �� �� �� �� Ժ"
Printer.PaintPicture Image1.Picture, 1, 1, 3, 3
Printer.PaintPicture Image1.Picture, 8, 1, 3, 3
Printer.FontSize = 12
Printer.CurrentX = 6
Printer.CurrentY = 4
Printer.ForeColor = vbBlack
Printer.Print Space(2) & "�� �� �� ¼ ��"
Printer.DrawStyle = 0  '��ʵ�ߴ�ӡ��VbDash 1 ���� VbDot 2����
                        'VbDashDot    3         �㻮��
                        'VbDashDotDot 4       ˫�㻮��
                        'VbInvisible  5           ����
                        'VbInsideSolid 6        ����ʵ��
 Printer.FontBold = False
Printer.Line (1, 5)-(17, 5)
Printer.Line (1, 5.1)-(17, 5.1)

Printer.CurrentX = 1
Printer.CurrentY = 6
Printer.Print Label4(0).ToolTipText & Label4(0).Caption & Space(2) & Label4(1).ToolTipText & Label4(1).Caption & Space(2) & Label4(2).ToolTipText & Label4(2).Caption
Printer.CurrentX = 1
Printer.CurrentY = 7
Printer.Print ; Label4(3).ToolTipText & Label4(3).Caption & Space(2) & Label4(4).ToolTipText & Label4(4).Caption & Space(2) & Label4(5).ToolTipText & Label4(5).Caption & Space(2) & Label4(6).ToolTipText & Label4(6).Caption
Printer.CurrentX = 1
Printer.CurrentY = 8
Printer.Line (1, 8.5)-(17, 8.5)

Printer.CurrentX = 1
Printer.CurrentY = 9

Printer.Print ; Label4(7).ToolTipText & Label4(7).Caption
'Printer.PSet (1, 10)
Printer.CurrentX = 1
Printer.CurrentY = 10
Printer.Print ; Label4(8).ToolTipText & Label4(8).Caption
'Printer.PSet (1, 13)
Printer.CurrentX = 1
Printer.CurrentY = 13
Printer.Print ; Label4(9).ToolTipText & Label4(9).Caption
Printer.CurrentX = 1
Printer.CurrentY = 15
Printer.Print ; Label4(10).ToolTipText & Label4(10).Caption
Printer.CurrentX = 1
Printer.CurrentY = 17
Printer.Print ; Label4(11).ToolTipText & Label4(11).Caption
Printer.CurrentX = 1
Printer.CurrentY = 22
Printer.Print "����ҽ��" & Label2.Caption
Printer.EndDoc
End Sub

Private Sub FlatScrollBar1_Change()
Pbox1.Move FlatScrollBar1.Value, FlatScrollBar2.Value
Label2.Caption = "X=" & FlatScrollBar1.Value
End Sub
Private Sub FlatScrollBar2_Change()
Pbox1.Move FlatScrollBar1.Value, -FlatScrollBar2.Value
Label3.Caption = "Y=" & FlatScrollBar2.Value
End Sub
Private Sub Form_Load()
Label2.Caption = ҽ������վMDI.StatusBar1.Panels(3).Text
Label5.Caption = ҽ������վMDI.StatusBar1.Panels(6).Text
Me.Left = 10000
End Sub

Private Sub Form_LostFocus()
MsgBox "���ȹرյ�ǰ���ڣ�"
End Sub

Private Sub Form_Resize()
Me.FlatScrollBar1.Top = Me.Height - 750
Me.FlatScrollBar2.Left = Me.Width - 500
End Sub


Private Sub Label11_Click()

End Sub

Private Sub Timer1_Timer()
'For i = 1 To 16
'For Y = 1 To 25
'Pbox1.Line (i, 1)-(i, 25)
'Pbox1.Line (1, Y)-(16, Y)
'Pbox1.CurrentX = i - 0.25
'Pbox1.CurrentY = 0.8
'Pbox1.Print i
'Pbox1.CurrentX = 0.8
'Pbox1.CurrentY = Y - 0.2
'Pbox1.Print Y
'Next Y
'Next i
Label1.Caption = Val(Label1.Caption)
Combo1.Text = (Label1.Caption * 100) & "%"
a = Label1.Caption
Pbox1.FontSize = 16
Pbox1.PSet (6, 3)
Pbox1.ForeColor = vbRed
Pbox1.FontBold = True
Pbox1.Print "�ĵ�������Ժ"
Pbox1.PaintPicture Image1.Picture, 1, 1, 3, 3
Pbox1.PaintPicture Image1.Picture, 10, 1, 3, 3
Pbox1.FontSize = 12
Pbox1.PSet (6, 4)
Pbox1.ForeColor = vbBlack
Pbox1.Print Space(2) & "�����¼��"
Pbox1.DrawStyle = 0   '��ʵ�ߴ�ӡ��VbDash 1 ���� VbDot 2����
                        'VbDashDot    3         �㻮��
                        'VbDashDotDot 4       ˫�㻮��
                        'VbInvisible  5           ����
                        'VbInsideSolid 6        ����ʵ��
 Pbox1.FontBold = False
Pbox1.Line (1, 5)-(17, 5)
Pbox1.Line (1, 5.1)-(17, 5.1)
Pbox1.PSet (1, 6)
Pbox1.Print Label4(0).ToolTipText & Label4(0).Caption & Space(2) & Label4(1).ToolTipText & Label4(1).Caption & Space(2) & Label4(2).ToolTipText & Label4(2).Caption
Pbox1.PSet (1, 7)
Pbox1.Print Label4(3).ToolTipText & Label4(3).Caption & Space(2) & Label4(4).ToolTipText & Label4(4).Caption & Space(2) & Label4(5).ToolTipText & Label4(5).Caption & Space(2) & Label4(6).ToolTipText & Label4(6).Caption
Pbox1.PSet (1, 8)
Pbox1.Line (1, 8.5)-(17, 8.5)
Pbox1.PSet (1, 9)
Pbox1.Print Label4(7).ToolTipText & Label4(7).Caption
Pbox1.PSet (1, 10)
Pbox1.Print Label4(8).ToolTipText & Label4(8).Caption
Pbox1.PSet (1, 13)
Pbox1.Print Label4(9).ToolTipText & Label4(9).Caption
Pbox1.PSet (1, 15)
Pbox1.Print Label4(10).ToolTipText & Label4(10).Caption
Pbox1.PSet (1, 17)
Pbox1.Print Label4(11).ToolTipText & Label4(11).Caption
Pbox1.PSet (1, 18)
Pbox1.Print
Pbox1.PSet (1, 19)
Pbox1.Print
Pbox1.PSet (1, 20)
Pbox1.Print "����ҽ����" & Label2.Caption
Pbox1.PSet (1, 21)
Pbox1.Print
Timer1.Interval = 0
End Sub
