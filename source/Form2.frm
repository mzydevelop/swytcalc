VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "��λһ���ۺ����ۼ�����"
   ClientHeight    =   7980
   ClientLeft      =   9150
   ClientTop       =   2505
   ClientWidth     =   14010
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7980
   ScaleWidth      =   14010
   Begin VB.CommandButton Command6 
      Caption         =   "�������°汾������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "���ձ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�л���У"
      BeginProperty Font 
         Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "һ���л�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ѧУ��ţ�"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "������λһ���ۺϷֳɼ���"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "��У��Ϣ"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "�鿴����ʱ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "���뱨��ϵͳ"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�鿴�����³�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�����ۺϷ�"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "�����㽭ʡ�߿��ɼ���"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�����ۺϲ��Գɼ���"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "���ڼ���..."
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����ѧУ��ţ�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��ѡ���ѧУ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "���ڼ���..."
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "����֧��QQ1172637796"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ӭ��!"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ӭ��ʹ����λһ���ۺ����ۼ�������"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Menu ���ڱ���� 
      Caption         =   "���ڱ����"
      Begin VB.Menu ��Ȩ��Ϣ 
         Caption         =   "��Ȩ��Ϣ"
      End
      Begin VB.Menu ���� 
         Caption         =   "��ϵ����"
      End
      Begin VB.Menu online 
         Caption         =   "�ٷ���վ"
      End
      Begin VB.Menu ������λһ���ۺ����ۼ����� 
         Caption         =   "������λһ���ۺ����ۼ�����"
      End
   End
   Begin VB.Menu ������� 
      Caption         =   "�������"
   End
   Begin VB.Menu ��Դ��ַ 
      Caption         =   "��Դ��վ"
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
   End
   Begin VB.Menu ע�� 
      Caption         =   "ע��"
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
        MsgBox "�ۺ����ʳɼ���߿��ɼ�����", vbCritical, "ϵͳ��Ϣ"
   Else
     swyts = xkf + zhf + gkcj * examb
     swyts = Int(10000 * swyts)
     swyts = swyts / 10000
   End If
Case 14
    If zhf > 120 Or gkcj > 1 / pc Then
        MsgBox "�ۺ����ʳɼ���߿��ɼ�����", vbCritical, "ϵͳ��Ϣ"
    Else
        zhf = zhf / 1.2
        swyts = xkf * xb + zhf * zb + gkcj * pc * examb * 100
        swyts = Int(10000 * swyts)
        swyts = swyts / 10000
    End If
Case 20
   If zhf > 100 Or gkcj > 1 / pc Then
        MsgBox "�ۺ����ʳɼ���߿��ɼ�����", vbCritical, "ϵͳ��Ϣ"
   Else
     swyts = xkf * xb + zhf * 7.5 * zb + gkcj * examb
     swyts = Int(10000 * swyts)
     swyts = swyts / 10000
   End If
Case 26
   If zhf > 100 Or gkcj > 1 / pc Then
        MsgBox "�ۺ����ʳɼ���߿��ɼ�����", vbCritical, "ϵͳ��Ϣ"
   Else
     swyts = xkf * xb + zhf * 7.5 * zb + gkcj * examb
     swyts = Int(10000 * swyts)
     swyts = swyts / 10000
   End If
Case Else
 
   If zhf > 100 Or gkcj > 1 / pc Then
   MsgBox "�ۺ����ʳɼ���߿��ɼ�����", vbCritical, "ϵͳ��Ϣ"
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
btime = btime + "����������ʱ�䡣"
MsgBox btime, vbInformation, "��������"
End Sub

Private Sub Command46_Click()
Form3.Show
End Sub

Private Sub Command5_Click()
xh = Text3.Text
If Len(xm) > 2 Then
Call sjk
Form2.Caption = "��λһ���ۺ����ۼ�����" + "(ѡ��ѧУ��" + xm + ")"
Timer1.Enabled = True
Else
MsgBox "������������ţ�1-39����", vbInformation, "ϵͳ��ʾ"
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
b = MsgBox("��ȷ��Ҫ�˳����û�ϵͳ��", vbInformation + vbYesNo, "��λһ���ۺ����ۼ�����")
If b = vbYes Then
Call setconfig
End
Else
MsgBox "��ȡ�����˳�ϵͳ��", vbOKOnly, "��λһ���ۺ����ۼ�����"
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
Label8.Caption = "��ӭ�μ�" + xm + "�ۺϷּ���!"
If Len(Label8.Caption) > 2 Then
Timer1.Enabled = False
Else
End If
End Sub











Private Sub ������λһ���ۺ����ۼ�����_Click()
frmAbout.Show
End Sub







Private Sub ��Դ��ַ_Click()

ShellExecute Me.hWnd, "open", git, "", "", 1
End Sub

Private Sub �������_Click()
MsgBox "����ʹ����鿴����������£���ǰ�汾Ϊbuild" + softverb, vbInformation, "��ǰ����汾��" + softverb + "��"
End Sub

Private Sub ��Ȩ��Ϣ_Click()
MsgBox "���������Ȩ���㽭ʡ2017��λһ�����������Ա��", vbInformation, "��Ȩ��Ϣ"
End Sub






Private Sub ע��_Click()
Unload Me
End Sub

Private Sub ����_Click()
MsgBox "QQ1172637796"
End Sub
