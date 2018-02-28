Attribute VB_Name = "rwini"
'****************************
'**INI�ļ���ȡд��ģ�����ͨ�� **
'****************************
'---------------------------------------------------------------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Function AppProfileName() As String
    
    AppProfileName = App.Path + "\config.ini" 'INI�ļ��洢��λ��
End Function

Public Function GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String) As String

    Dim ResultString As String * 144, Temp As Integer
    Dim s As String, i As Integer
    Temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, AppProfileName())
    '�����ؼ��ʵ�ֵ
    If Temp% > 0 Then '�ؼ��ʵ�ֵ��Ϊ��
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
        '��ȱʡֵд��INI�ļ�
        s = DefString
    End If
    GetIniS = s

End Function

Public Function GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Integer) As Integer

    Dim d As Long, s As String
    d = DefValue 'DefValueΪ�ؼ��ʵ�ȱʡֵ
    GetIniN = GetPrivateProfileInt(SectionName, KeyWord, DefValue, AppProfileName())
    If d <> DefValue Then
        s = "" & d
        d = WritePrivateProfileString(SectionName, KeyWord, s, AppProfileName())
    End If
    
End Function

Public Sub SetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String)

    Dim res%
    res% = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProfileName()) 'ValStrΪҪд��ini�ļ��Ĺؼ��ʵ�ֵ
End Sub

Public Sub SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Integer)

    Dim res%, s$
    s$ = Str$(ValInt)
    res% = WritePrivateProfileString(SectionName, KeyWord, s$, AppProfileName()) 'ValIntΪҪд��ini�ļ��Ĺؼ��ʵ�ֵ
End Sub

