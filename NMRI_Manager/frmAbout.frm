VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于 NMRI_Manager"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "关于 NMRI_Manager"
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Tag             =   "确定"
      Top             =   2625
      Width           =   1467
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "系统信息(&S)..."
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Tag             =   "系统信息(&S)..."
      Top             =   3075
      Width           =   1452
   End
   Begin VB.Label lblDescription 
      Caption         =   "应用程序描述"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   6
      Tag             =   "应用程序描述"
      Top             =   1125
      Width           =   4092
   End
   Begin VB.Label lblTitle 
      Caption         =   "应用程序标题"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Tag             =   "应用程序标题"
      Top             =   240
      Width           =   4092
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   225
      X2              =   5657
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5657
      Y1              =   2445
      Y2              =   2445
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Tag             =   "版本"
      Top             =   780
      Width           =   4092
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "警告：本程序受Apache Version 2.0开源协议保护，详细内容参见http://www.apache.org/licenses/LICENSE-2.0"
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   3
      Tag             =   "警告: ..."
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 注册键安全选项...
Const KEY_ALL_ACCESS = &H2003F
                                          

' 注册键根类型...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode 空结尾字符串
Const REG_DWORD = 4                      ' 32位数


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub



Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub


Private Sub cmdOK_Click()
        Unload Me
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' 从注册表获得系统信息程序路径\名称...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' 仅从注册表获得系统信息程序路径...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' 验证已知的 32 位文件版本的存在
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' 错误 - 文件找不到...
                Else
                        GoTo SysInfoErr
                End If
        ' 错误 - 注册表项找不到...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "此时系统信息不可用", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' 循环记数器
        Dim rc As Long                                          ' 返回代码
        Dim hKey As Long                                        ' 打开的注册表键句柄
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' 注册表键数据类型
        Dim tmpVal As String                                    ' 临时存储一个注册表键值
        Dim KeyValSize As Long                                  ' 注册表键变量大小
        '------------------------------------------------------------
        ' 在键根{HKEY_LOCAL_MACHINE...}之下打开注册键
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表键
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 错误处理...
        

        tmpVal = String$(1024, 0)                             ' 分配变量空间
        KeyValSize = 1024                                       ' 标记变量大小
        

        '------------------------------------------------------------
        ' 检索注册表键值...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' 获得/创建键值
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 错误处理
      

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)


        '------------------------------------------------------------
        ' 决定转换的键值类型...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' 搜索数据类型...
        Case REG_SZ                                             ' 字符串注册表键数据类型
                KeyVal = tmpVal                                     ' 复制字符串值
        Case REG_DWORD                                          ' 双精度注册表键数据类型
                For i = Len(tmpVal) To 1 Step -1                    ' 转换每一页
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 一个字符一个字符地生成值
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' 转换双精度为字符串
        End Select
        

        GetKeyValue = True                                      ' 返回成功
        rc = RegCloseKey(hKey)                                  ' 关闭注册表键
        Exit Function                                           ' 退出
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' 设返回值为空字符串
        GetKeyValue = False                                     ' 返回失败
        rc = RegCloseKey(hKey)                                  ' 关闭注册表键
End Function

