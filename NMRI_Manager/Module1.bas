Attribute VB_Name = "Module1"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public fMainForm As frmMain


Sub Main()
    Dim fLogin As New frmLogin
    'Dim fTip As New frmTip
    Dim ShowAtStartup As Long
    fLogin.Show vbModal
    If Not fLogin.OK Then
        '登录失败，退出应用程序
        End
    End If
    Unload fLogin
    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Sleep 500
    Load fMainForm
    Unload frmSplash
    fMainForm.Show
    ' 察看在启动时是否将被显示
    ShowAtStartup = GetSetting(App.EXEName, "Options", "在启动时显示提示", 1)
    If ShowAtStartup = 1 Then
    frmTip.Show vbModal, fMainForm
    End If
End Sub

' 读取INI文件函数
Public Function ReadINI(ByVal sSection As String, ByVal sKey As String, ByVal sDefault As String, ByVal sFileName As String) As String
    Dim sRetVal As String
    Dim nBufferSize As Long
    nBufferSize = 255
    sRetVal = String$(nBufferSize, vbNullChar)
    GetPrivateProfileString sSection, sKey, sDefault, sRetVal, nBufferSize, sFileName
    ReadINI = Left$(sRetVal, InStr(sRetVal, vbNullChar) - 1)
End Function

' 写入INI文件函数
Public Function WriteINI(ByVal sSection As String, ByVal sKey As String, ByVal sValue As String, ByVal sFileName As String) As Boolean
    WriteINI = WritePrivateProfileString(sSection, sKey, sValue, sFileName)
End Function

