VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "关于三位一体综合评价计算器"
   ClientHeight    =   4785
   ClientLeft      =   2415
   ClientTop       =   2010
   ClientWidth     =   8580
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3302.692
   ScaleMode       =   0  'User
   ScaleWidth      =   8057.063
   Begin VB.CommandButton Command1 
      Caption         =   "点击访问GitHub项目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   850
      Left            =   4200
      Top             =   240
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   420
      Left            =   240
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   252.84
      ScaleMode       =   0  'User
      ScaleWidth      =   252.84
      TabIndex        =   0
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "正在加载中..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "应用程序描述:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1050
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   7845
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "三位一体综合评价计算器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   2925
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "版本 v1.0.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1080
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":194C
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   900
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   8295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
ShellExecute Me.hWnd, "open", git, "", "", 1
End Sub

Private Sub Form_Load()
    App.Title = "三位一体综合评价计算器"
    Me.Caption = "关于 " & App.Title
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    Timer1.Enabled = True
End Sub





Private Sub Timer1_Timer()
lblVersion.Caption = "版本 v" + softverm + "." + softverv
Label1.Caption = "Build" + softverb
lblDescription.Caption = softabout
If Len(Label1.Caption) > 3 Then
Timer1.Enabled = False
Else
End If
End Sub
