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



Public Sub loaddata()
'softverm = GetIniS("softinfo", "���汾��", "")
'softverv = GetIniS("softinfo", "���汾��", "")
'softverb = GetIniS("softinfo", "Build", "")
'snumber = GetIniS("copyright", "���к�", "")
'softabout = GetIniS("softinfo", "�������", "")
If softverm = "" Then
softverm = "1.0"
softverv = "1"
softverb = "0014"
snumber = "���к��ݲ��ṩ��ȫ����Ȩ��"
softabout = "�����˱���ѧ����Ϣ�������ļ����ܣ�Ŀǰ�����С����汾���µ�Build0014��"
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
End Sub
