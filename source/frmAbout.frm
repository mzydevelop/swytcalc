VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "������λһ���ۺ����ۼ�����"
   ClientHeight    =   3555
   ClientLeft      =   2415
   ClientTop       =   2010
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   Begin VB.Timer Timer1 
      Interval        =   850
      Left            =   120
      Top             =   3000
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
      TabIndex        =   1
      Top             =   240
      Width           =   420
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   345
      Left            =   4125
      TabIndex        =   0
      Top             =   2625
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "���ĵ�����Ϣ"
      Height          =   345
      Left            =   4140
      TabIndex        =   2
      Top             =   3075
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "���ڼ�����..."
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
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ�ó�������:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "��λһ���ۺ����ۼ�����"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   5
      Top             =   240
      Width           =   2925
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "�汾 v1.0.0"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   6
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "����: ��Ȩ���У�����ؾ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   4
      Top             =   2625
      Width           =   3630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ע���ؼ��ְ�ȫѡ��...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ע���ؼ��� ROOT ����...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' �����Ŀյ��ս��ַ���
Const REG_DWORD = 4                      ' 32λ����

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    App.Title = "��λһ�������"
    Me.Caption = "���� " & App.Title
    lblVersion.Caption = "�汾 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    Timer1.Enabled = True
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ��ͼ��ע����л��ϵͳ��Ϣ�����·��������...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' ��ͼ����ע����л��ϵͳ��Ϣ�����·��...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' ��֪32λ�ļ��汾����Чλ��
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' ���� - �ļ����ܱ��ҵ�...
        Else
            GoTo SysInfoErr
        End If
    ' ���� - ע�����Ӧ��Ŀ���ܱ��ҵ�...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "��ʱϵͳ��Ϣ������", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' ѭ��������
    Dim rc As Long                                          ' ���ش���
    Dim hKey As Long                                        ' �򿪵�ע���ؼ��־��
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ע���ؼ�����������
    Dim tmpVal As String                                    ' ע���ؼ���ֵ����ʱ�洢��
    Dim KeyValSize As Long                                  ' ע���ؼ��Ա����ĳߴ�
    '------------------------------------------------------------
    ' �� {HKEY_LOCAL_MACHINE...} �µ� RegKey
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע���ؼ���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ����ߴ�
    
    '------------------------------------------------------------
    ' ����ע���ؼ��ֵ�ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ���/�����ؼ���ֵ
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 ��ӳ�����ս��ַ���...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null ���ҵ�,���ַ����з������
    Else                                                    ' WinNT û�п��ս��ַ���...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null û�б��ҵ�, �����ַ���
    End If
    '------------------------------------------------------------
    ' ����ת���Ĺؼ��ֵ�ֵ����...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' ������������...
    Case REG_SZ                                             ' �ַ���ע��ؼ�����������
        KeyVal = tmpVal                                     ' �����ַ�����ֵ
    Case REG_DWORD                                          ' ���ֽڵ�ע���ؼ�����������
        For i = Len(tmpVal) To 1 Step -1                    ' ��ÿλ����ת��
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' ����ֵ�ַ��� By Char��
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' ת�����ֽڵ��ַ�Ϊ�ַ���
    End Select
    
    GetKeyValue = True                                      ' ���سɹ�
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
    Exit Function                                           ' �˳�
    
GetKeyError:      ' �������������...
    KeyVal = ""                                             ' ���÷���ֵ�����ַ���
    GetKeyValue = False                                     ' ����ʧ��
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
End Function

Private Sub Timer1_Timer()
lblVersion.Caption = "�汾 v" + softverm + "." + softverv
Label1.Caption = "Build" + softverb
lblDescription.Caption = softabout
If Len(Label1.Caption) > 3 Then
Timer1.Enabled = False
Else
End If
End Sub
