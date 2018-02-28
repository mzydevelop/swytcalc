Attribute VB_Name = "rwini"
'****************************
'**INI文件读取写入模块代码通用 **
'****************************
'---------------------------------------------------------------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Function AppProfileName() As String
    
    AppProfileName = App.Path + "\config.ini" 'INI文件存储的位置
End Function

Public Function GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String) As String

    Dim ResultString As String * 144, Temp As Integer
    Dim s As String, i As Integer
    Temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, AppProfileName())
    '检索关键词的值
    If Temp% > 0 Then '关键词的值不为空
        s = ""
        For i = 1 To 144
            If Asc(Mid$(ResultString, i, 1)) = 0 Then
                Exit For
            Else
                s = s & Mid$(ResultString, i, 1)
            End If
        Next
    Else
        Temp% = WritePrivateProfileString(SectionName, KeyWord, DefString, AppProfileName())
        '将缺省值写入INI文件
        s = DefString
    End If
    GetIniS = s

End Function

Public Function GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Integer) As Integer

    Dim d As Long, s As String
    d = DefValue 'DefValue为关键词的缺省值
    GetIniN = GetPrivateProfileInt(SectionName, KeyWord, DefValue, AppProfileName())
    If d <> DefValue Then
        s = "" & d
        d = WritePrivateProfileString(SectionName, KeyWord, s, AppProfileName())
    End If
    
End Function

Public Sub SetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String)

    Dim res%
    res% = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProfileName()) 'ValStr为要写入ini文件的关键词的值
End Sub

Public Sub SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Integer)

    Dim res%, s$
    s$ = Str$(ValInt)
    res% = WritePrivateProfileString(SectionName, KeyWord, s$, AppProfileName()) 'ValInt为要写入ini文件的关键词的值
End Sub

