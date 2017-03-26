VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "三位一体综合评价计算器"
   ClientHeight    =   7980
   ClientLeft      =   9150
   ClientTop       =   2505
   ClientWidth     =   14010
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7980
   ScaleWidth      =   14010
   Begin VB.CommandButton Command6 
      Caption         =   "下载最新版本计算器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10320
      TabIndex        =   24
      Top             =   6000
      Width           =   3495
   End
   Begin VB.CommandButton Command46 
      Caption         =   "对照表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      TabIndex        =   23
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      Caption         =   "切换高校"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   10200
      TabIndex        =   19
      Top             =   120
      Width           =   3615
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "一键切换"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         MaskColor       =   &H80000010&
         TabIndex        =   21
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label12 
         Caption         =   "学校序号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "计算结果"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   11
      Top             =   4680
      Width           =   9615
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "【分数】"
         BeginProperty Font 
            Name            =   "隶书"
            Size            =   36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   4200
         TabIndex        =   13
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "您的三位一体综合分成绩："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   680
      Left            =   8520
      Top             =   360
   End
   Begin VB.Frame Frame1 
      Caption         =   "高校信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   9615
      Begin VB.CommandButton Command4 
         Caption         =   "查看报名时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4920
         TabIndex        =   18
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "进入报名系统"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   17
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "查看招生章程"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "计算综合分"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   5880
         TabIndex        =   15
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   8
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "您的浙江省高考成绩："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "您的综合测试成绩："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "正在加载..."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   5
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "您的学校序号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "您选择的学校："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "正在加载..."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   2
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "技术支持QQ1172637796"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10440
      TabIndex        =   14
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎您!"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   6960
      Width           =   8415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎您使用三位一体综合评价计算器！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
   Begin VB.Menu 关于本软件 
      Caption         =   "关于本软件"
      Begin VB.Menu 授权信息 
         Caption         =   "授权信息"
      End
      Begin VB.Menu 作者 
         Caption         =   "联系作者"
      End
      Begin VB.Menu online 
         Caption         =   "官方网站"
      End
      Begin VB.Menu 关于三位一体综合评价计算器 
         Caption         =   "关于三位一体综合评价计算器"
      End
   End
   Begin VB.Menu 软件更新 
      Caption         =   "软件更新"
   End
   Begin VB.Menu 开源网址 
      Caption         =   "开源网站"
   End
   Begin VB.Menu 帮助 
      Caption         =   "帮助"
   End
   Begin VB.Menu 注销 
      Caption         =   "注销"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command8_Click()
frmAbout.Show
End Sub

Private Sub Command1_Click()
zhf = Val(Text1.Text)
gkcj = Val(Text2.Text)
Select Case xh
Case 6
   If zhf > 225 Or gkcj > 1 / pc Then
        MsgBox "综合素质成绩或高考成绩有误", vbCritical, "系统消息"
   Else
     swyts = xkf + zhf + gkcj * examb
     swyts = Int(10000 * swyts)
     swyts = swyts / 10000
   End If
Case 14
    If zhf > 120 Or gkcj > 1 / pc Then
        MsgBox "综合素质成绩或高考成绩有误", vbCritical, "系统消息"
    Else
        zhf = zhf / 1.2
        swyts = xkf * xb + zhf * zb + gkcj * pc * examb * 100
        swyts = Int(10000 * swyts)
        swyts = swyts / 10000
    End If
Case 20
   If zhf > 100 Or gkcj > 1 / pc Then
        MsgBox "综合素质成绩或高考成绩有误", vbCritical, "系统消息"
   Else
     swyts = xkf * xb + zhf * 7.5 * zb + gkcj * examb
     swyts = Int(10000 * swyts)
     swyts = swyts / 10000
   End If
Case 26
   If zhf > 100 Or gkcj > 1 / pc Then
        MsgBox "综合素质成绩或高考成绩有误", vbCritical, "系统消息"
   Else
     swyts = xkf * xb + zhf * 7.5 * zb + gkcj * examb
     swyts = Int(10000 * swyts)
     swyts = swyts / 10000
   End If
Case Else
 
   If zhf > 100 Or gkcj > 1 / pc Then
   MsgBox "综合素质成绩或高考成绩有误", vbCritical, "系统消息"
   Else
      swyts = xkf * xb + zhf * zb + gkcj * pc * examb * 100
      swyts = Int(10000 * swyts)
      swyts = swyts / 10000

   End If
End Select
Label10.Caption = Str(swyts)
End Sub

Private Sub Command2_Click()
ShellExecute Me.hWnd, "open", wz, "", "", 1
End Sub

Private Sub Command3_Click()
ShellExecute Me.hWnd, "open", bm, "", "", 1
End Sub

Private Sub Command4_Click()
btime = btime + "请勿错过报名时间。"
MsgBox btime, vbInformation, "友情提醒"
End Sub

Private Sub Command46_Click()
Form3.Show
End Sub

Private Sub Command5_Click()
xh = Text3.Text
If Len(xm) > 2 Then
Call sjk
Form2.Caption = "三位一体综合评价计算器" + "(选择学校：" + xm + ")"
Timer1.Enabled = True
Else
MsgBox "请重新输入序号（1-39）！", vbInformation, "系统提示"
End If
End Sub

Private Sub Command6_Click()
newv = "http://mzy115.is-programmer.com/user_files/mzy115/File/soft/newjsq.rar"
ShellExecute Me.hWnd, "open", newv, "", "", 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Load()
Timer1.Enabled = True

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim b As String
b = MsgBox("你确定要退出本用户系统吗？", vbInformation + vbYesNo, "三位一体综合评价计算器")
If b = vbYes Then
Call setconfig
End
Else
MsgBox "你取消了退出系统！", vbOKOnly, "三位一体综合评价计算器"
Cancel = -1
End If
End Sub



Private Sub online_Click()
gw = "http://mzy115.is-programmer.com/2017/2/18/swyt2017.208634.html"
ShellExecute Me.hWnd, "open", gw, "", "", 1
End Sub

Private Sub Timer1_Timer()

Label3.Caption = xm
Label5.Caption = xh
Label8.Caption = "欢迎参加" + xm + "综合分计算!"
If Len(Label8.Caption) > 2 Then
Timer1.Enabled = False
Else
End If
End Sub











Private Sub 关于三位一体综合评价计算器_Click()
frmAbout.Show
End Sub







Private Sub 开源网址_Click()

ShellExecute Me.hWnd, "open", git, "", "", 1
End Sub

Private Sub 软件更新_Click()
MsgBox "请访问官网查看最新软件更新！当前版本为build" + softverb, vbInformation, "当前软件版本【" + softverb + "】"
End Sub

Private Sub 授权信息_Click()
MsgBox "本软件已授权给浙江省2017三位一体所有相关人员！", vbInformation, "授权信息"
End Sub






Private Sub 注销_Click()
Unload Me
End Sub

Private Sub 作者_Click()
MsgBox "QQ1172637796"
End Sub
