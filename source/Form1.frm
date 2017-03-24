VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "2017浙江三位一体综合评价计算器 Made By MZY v1.00 Build0010"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12405
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   12405
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command46 
      Caption         =   "对照表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   78
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command45 
      Caption         =   "点此打开招生章程"
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
      Left            =   9240
      TabIndex        =   77
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command16 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   67
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command36 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   51
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command35 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   50
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command34 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   49
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command33 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   48
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command32 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   47
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command31 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   46
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command30 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   45
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command29 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   44
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command28 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   43
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command26 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   42
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command25 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   41
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command24 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   40
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command23 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   39
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command22 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   38
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command21 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   37
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command20 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   36
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command19 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   35
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command18 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   34
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   33
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   32
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   31
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   29
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   28
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   26
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   24
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   23
      Top             =   3120
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "学业水平考试信息"
      Height          =   3615
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   11895
      Begin VB.CommandButton Command44 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10560
         TabIndex        =   75
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command43 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10560
         TabIndex        =   74
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton Command42 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10560
         TabIndex        =   73
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command41 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10560
         TabIndex        =   72
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command40 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   71
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command39 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   70
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command38 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   69
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command27 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   68
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   66
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command37 
         Caption         =   "【学考信息区】"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   65
         Top             =   2880
         Width           =   6735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "保存学考信息"
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
         Left            =   5880
         TabIndex        =   21
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "您的学考成绩="
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
         Left            =   120
         TabIndex        =   64
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label108 
         BackStyle       =   0  'Transparent
         Caption         =   "未选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   63
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label109 
         BackStyle       =   0  'Transparent
         Caption         =   "未选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   62
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label110 
         BackStyle       =   0  'Transparent
         Caption         =   "未选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   61
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label101 
         BackStyle       =   0  'Transparent
         Caption         =   "未选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   60
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label102 
         BackStyle       =   0  'Transparent
         Caption         =   "未选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   59
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label103 
         BackStyle       =   0  'Transparent
         Caption         =   "未选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   58
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label106 
         BackStyle       =   0  'Transparent
         Caption         =   "未选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   57
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label105 
         BackStyle       =   0  'Transparent
         Caption         =   "未选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   56
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label104 
         BackStyle       =   0  'Transparent
         Caption         =   "未选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   55
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label107 
         BackStyle       =   0  'Transparent
         Caption         =   "未选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   52
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "语文"
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
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "数学"
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
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "英语"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "政治"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "历史"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "地理"
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
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "物理"
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
         Left            =   6000
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "化学"
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
         Left            =   6000
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "生物"
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
         Left            =   6000
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "技术"
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
         Left            =   6000
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1200
      Left            =   6960
      Top             =   7200
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "开始计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      MaskColor       =   &H80000012&
      TabIndex        =   7
      Top             =   5400
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询学校参数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   4
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   3735
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
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   $"Form1.frx":038A
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7560
      TabIndex        =   76
      Top             =   5520
      Width           =   4455
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "未选择"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   54
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "未选择"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   53
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   7200
      Width           =   5055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "当前状态："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "输入（1-39，学校序号参见对照表。）"
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
      Left            =   3960
      TabIndex        =   6
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Made By MZY Develop 20170318"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   6600
      Width           =   6615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "学校名:"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "学校序号："
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
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
xh = Text1.Text
If xkbc = False Then
MsgBox "请先保存学考信息！", vbExclamation, "系统信息"
Else
Call sjk
Text2.Text = xm
End If
End Sub

Private Sub Command10_Click()
xkb(4) = "A"
Label104.Caption = xkb(4)
xk(4) = 1
End Sub

Private Sub Command11_Click()
xkb(8) = "A"
Label108.Caption = xkb(8)
xk(8) = 1
End Sub

Private Sub Command12_Click()
xkb(7) = "A"
Label107.Caption = xkb(7)
xk(7) = 1
End Sub

Private Sub Command13_Click()
xkb(3) = "A"
Label103.Caption = xkb(3)
xk(3) = 1
End Sub

Private Sub Command14_Click()
xkb(2) = "A"
Label102.Caption = xkb(2)
xk(2) = 1
End Sub

Private Sub Command15_Click()
xkb(1) = "B"
Label101.Caption = xkb(1)
xk(1) = 1
End Sub



Private Sub Command16_Click()
xkb(2) = "D"
Label102.Caption = xkb(2)
xk(2) = 1
End Sub

Private Sub Command17_Click()
xkb(10) = "B"
Label110.Caption = xkb(10)
xk(10) = 1
End Sub

Private Sub Command18_Click()
xkb(9) = "B"
Label109.Caption = xkb(9)
xk(9) = 1
End Sub

Private Sub Command19_Click()
xkb(8) = "B"
Label108.Caption = xkb(8)
xk(8) = 1
End Sub

Private Sub Command2_Click()
If Len(xm) >= 2 Then
Label6.Caption = "正在载入" + xm + "参数"
Timer1.Enabled = True
Else
Text1.Text = ""
MsgBox "请先输入学校序号查询学校参数后再登陆，序号为（1-39）！", vbInformation, "系统消息"
End If
End Sub

Private Sub Command20_Click()
xkb(7) = "B"
Label107.Caption = xkb(7)
xk(7) = 1
End Sub

Private Sub Command21_Click()
xkb(6) = "B"
Label106.Caption = xkb(6)
xk(6) = 1
End Sub

Private Sub Command22_Click()
xkb(5) = "B"
Label105.Caption = xkb(5)
xk(5) = 1
End Sub

Private Sub Command23_Click()
xkb(4) = "B"
Label104.Caption = xkb(4)
xk(4) = 1
End Sub

Private Sub Command24_Click()
xkb(3) = "B"
Label103.Caption = xkb(3)
xk(3) = 1
End Sub

Private Sub Command25_Click()
xkb(2) = "B"
Label102.Caption = xkb(2)
xk(2) = 1
End Sub

Private Sub Command26_Click()
xkb(1) = "C"
Label101.Caption = xkb(1)
xk(1) = 1
End Sub



Private Sub Command27_Click()
xkb(3) = "D"
Label103.Caption = xkb(3)
xk(3) = 1
End Sub

Private Sub Command28_Click()
xkb(10) = "C"
Label110.Caption = xkb(10)
xk(10) = 1
End Sub

Private Sub Command29_Click()
xkb(9) = "C"
Label109.Caption = xkb(9)
xk(9) = 1
End Sub

Private Sub Command3_Click()
Dim xkq As Integer
Dim p As Integer
xkq = xk(1) + xk(2) + xk(3) + xk(4) + xk(5) + xk(6) + xk(7) + xk(8) + xk(9) + xk(10)
If xkq < 10 Then
MsgBox "你有未填写的学考信息，保存失败！", vbOKOnly, "三位一体综合评价计算器"
Else
  For p = 1 To 10
  If xkb(p) = "A" Then
  ag = ag + 1
  Else
    If xkb(p) = "B" Then
    bg = bg + 1
    Else
      If xkb(p) = "C" Then
      cg = cg + 1
      Else
        If xkb(p) = "D" Then
        dg = dg + 1
        End If
      End If
    End If
  End If
  Next p
  eg = 10 - (ag + bg + cg + dg)
  Command37.Caption = "您的A有" + Str(ag) + "个，B有" + Str(bg) + "个，C有" + Str(cg) + "个,D有" + Str(dg) + "个。"

  xkbc = True
  MsgBox "学考信息保存成功！", vbOKOnly, "三位一体综合评价计算器"
  Command3.Enabled = False
End If
End Sub

Private Sub Command30_Click()
xkb(6) = "C"
Label106.Caption = xkb(6)
xk(6) = 1
End Sub

Private Sub Command31_Click()
xkb(5) = "C"
Label105.Caption = xkb(5)
xk(5) = 1
End Sub

Private Sub Command32_Click()
xkb(4) = "C"
Label104.Caption = xkb(4)
xk(4) = 1
End Sub

Private Sub Command33_Click()
xkb(3) = "C"
Label103.Caption = xkb(3)
xk(3) = 1
End Sub

Private Sub Command34_Click()
xkb(8) = "C"
Label108.Caption = xkb(8)
xk(8) = 1
End Sub

Private Sub Command35_Click()
xkb(7) = "C"
Label107.Caption = xkb(7)
xk(7) = 1
End Sub

Private Sub Command36_Click()
xkb(2) = "C"
Label102.Caption = xkb(2)
xk(2) = 1
End Sub

Private Sub Command38_Click()
xkb(4) = "D"
Label104.Caption = xkb(4)
xk(4) = 1
End Sub

Private Sub Command39_Click()
xkb(5) = "D"
Label105.Caption = xkb(5)
xk(5) = 1
End Sub

Private Sub Command4_Click()
xkb(1) = "A"
Label101.Caption = xkb(1)
xk(1) = 1
End Sub



Private Sub Command40_Click()
xkb(6) = "D"
Label106.Caption = xkb(6)
xk(6) = 1
End Sub

Private Sub Command41_Click()
xkb(7) = "D"
Label107.Caption = xkb(7)
xk(7) = 1
End Sub

Private Sub Command42_Click()
xkb(8) = "D"
Label108.Caption = xkb(8)
xk(8) = 1
End Sub

Private Sub Command43_Click()
xkb(9) = "D"
Label109.Caption = xkb(9)
xk(9) = 1
End Sub

Private Sub Command44_Click()
xkb(10) = "D"
Label110.Caption = xkb(10)
xk(10) = 1
End Sub

Private Sub Command45_Click()
ShellExecute Me.hWnd, "open", wz, "", "", 1
End Sub

Private Sub Command46_Click()
Form3.Show
End Sub

Private Sub Command5_Click()
xkb(1) = "D"
Label101.Caption = xkb(1)
xk(1) = 1
End Sub

Private Sub Command6_Click()
xkb(10) = "A"
Label110.Caption = xkb(10)
xk(10) = 1
End Sub

Private Sub Command7_Click()
xkb(9) = "A"
Label109.Caption = xkb(9)
xk(9) = 1
End Sub

Private Sub Command8_Click()
xkb(6) = "A"
Label106.Caption = xkb(6)
xk(6) = 1
End Sub

Private Sub Command9_Click()
xkb(5) = "A"
Label105.Caption = xkb(5)
xk(5) = 1
End Sub

Private Sub Form_Load()
Dim a As Integer
Label6.Caption = "计算器登陆界面"
Timer1.Enabled = False
ag = 0
bg = 0
cg = 0
dg = 0
eg = 0
pc = 1 / 750

gly = False
xkbc = False
For a = 1 To 10
xk(a) = 0
Next a
Call loaddata
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim b As String
b = MsgBox("你确定要退出登陆吗？", vbInformation + vbYesNo, "三位一体综合评价计算器")
If b = vbYes Then
Call setconfig
End
Else
MsgBox "你取消了退出！", vbOKOnly, "三位一体综合评价计算器"
Cancel = -1
End If
End Sub


Private Sub Label20_Click()

End Sub

Private Sub Timer1_Timer()
If Len(xm) >= 2 Then
Me.Hide
Form2.Show
Form2.Caption = "三位一体综合评价计算器" + "版本：build" + softverb + "(选择学校：" + xm + ")"
Timer1.Enabled = False
Else
End If
End Sub
