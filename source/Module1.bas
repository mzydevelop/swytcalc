Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public xh As String
Public xm As String
Public rq As String
Public sj As String
Public sysver As String
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
    (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
    ByVal hIcon As Long) As Long
Public softverm As String
Public softverv As String
Public softverb As String
Public snumber As String
Public softabout As String
Public j As Integer
Public i As Integer
Public ag As Integer
Public bg As Integer
Public cg As Integer
Public dg As Integer
Public eg As Integer
Public xkbc As Boolean
Public xk(1 To 10) As Integer
Public xkb(1 To 10) As String
Public xb As Double
Public zb As Double
Public examb As Double
Public xkf As Double
Public cxk As Double
Public gkcj As Double
Public zhf As Double
Public swyts As Double
Public pc As Double
Public gw As String
Public wz As String
Public bm As String
Public btime As String
Public cs As String
Public newv As String
Public git As String
Public savekey As String
Public dq As Boolean
Public Sub loadxk()
savekey = GetIniS("xk", "savekey", "")
If savekey = "" Then
dq = False
MsgBox "��δ���������", vbInformation, "���ö�ȡ��Ϣ"
Else
dq = True
xkb(1) = GetIniS("xk", "����", "")
xkb(2) = GetIniS("xk", "��ѧ", "")
xkb(3) = GetIniS("xk", "Ӣ��", "")
xkb(4) = GetIniS("xk", "����", "")
xkb(5) = GetIniS("xk", "��ʷ", "")
xkb(6) = GetIniS("xk", "����", "")
xkb(7) = GetIniS("xk", "����", "")
xkb(8) = GetIniS("xk", "��ѧ", "")
xkb(9) = GetIniS("xk", "����", "")
xkb(10) = GetIniS("xk", "����", "")
MsgBox "���ģ�" + xkb(1) + "    ��ѧ��" + xkb(2) + "    Ӣ�" + xkb(3) + "    ���Σ�" + xkb(4) + "    ��ʷ��" + xkb(5) + "    ����" + xkb(6) + "    ����" + xkb(7) + "    ��ѧ��" + xkb(8) + "  ���" + xkb(9) + "    ������" + xkb(10), vbQuestion, "ѧ����Ϣ��ȡ���"

Dim p As Integer


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

End If
End Sub


Public Sub loaddata()
'softverm = GetIniS("softinfo", "���汾��", "")
'softverv = GetIniS("softinfo", "���汾��", "")
'softverb = GetIniS("softinfo", "Build", "")
'snumber = GetIniS("copyright", "���к�", "")
'softabout = GetIniS("softinfo", "�������", "")
If softverm = "" Then
softverm = "1.0"
softverv = "2"
softverb = "0018"
snumber = "���к��ݲ��ṩ��ȫ����Ȩ��"
softabout = "�����˴������ļ���ȡѧ����Ϣ���ܣ��޸��ദBUG��Ŀǰ�����С����汾���µ�Build0018��"
End If

End Sub
Public Sub setconfig()
SetIniS "softinfo", "���汾��", softverm   '��ȡ���õĲ���
SetIniS "softinfo", "���汾��", softverv
SetIniS "softinfo", "Build", softverb
SetIniS "copyright", "���к�", snumber
SetIniS "softinfo", "�������", softabout
End Sub

Public Sub xkout()
savekey = Str(Date) + Str(Time) + "@" + Str(softverm) + Str(softverv) + Str(softverb)
SetIniS "xk", "����", xkb(1)    '��ȡ���õĲ���
SetIniS "xk", "��ѧ", xkb(2)
SetIniS "xk", "Ӣ��", xkb(3)
SetIniS "xk", "����", xkb(4)
SetIniS "xk", "��ʷ", xkb(5)
SetIniS "xk", "����", xkb(6)
SetIniS "xk", "����", xkb(7)
SetIniS "xk", "��ѧ", xkb(8)
SetIniS "xk", "����", xkb(9)    '��ȡ���õĲ���
SetIniS "xk", "����", xkb(10)
SetIniS "xk", "savekey", savekey
End Sub
