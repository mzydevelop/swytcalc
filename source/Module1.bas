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
'softverm = GetIniS("softinfo", "主版本号", "")
'softverv = GetIniS("softinfo", "副版本号", "")
'softverb = GetIniS("softinfo", "Build", "")
'snumber = GetIniS("copyright", "序列号", "")
'softabout = GetIniS("softinfo", "软件描述", "")
If softverm = "" Then
softverm = "1.0"
softverv = "1"
softverb = "0014"
snumber = "序列号暂不提供（全部授权）"
softabout = "增加了保存学考信息到配置文件功能，目前测试中。，版本更新到Build0014。"
End If

End Sub
Public Sub setconfig()
SetIniS "softinfo", "主版本号", softverm   '获取设置的参数
SetIniS "softinfo", "副版本号", softverv
SetIniS "softinfo", "Build", softverb
SetIniS "copyright", "序列号", snumber
SetIniS "softinfo", "软件描述", softabout
End Sub

Public Sub xkout()
SetIniS "xk", "语文", xkb(1)    '获取设置的参数
SetIniS "xk", "数学", xkb(2)
SetIniS "xk", "英语", xkb(3)
SetIniS "xk", "政治", xkb(4)
SetIniS "xk", "历史", xkb(5)
SetIniS "xk", "地理", xkb(6)
SetIniS "xk", "物理", xkb(7)
SetIniS "xk", "化学", xkb(8)
SetIniS "xk", "生物", xkb(9)    '获取设置的参数
SetIniS "xk", "技术", xkb(10)
End Sub
