VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���� NMRI_Manager"
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
   StartUpPosition =   1  '����������
   Tag             =   "���� NMRI_Manager"
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
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Tag             =   "ȷ��"
      Top             =   2625
      Width           =   1467
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ϵͳ��Ϣ(&S)..."
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Tag             =   "ϵͳ��Ϣ(&S)..."
      Top             =   3075
      Width           =   1452
   End
   Begin VB.Label lblDescription 
      Caption         =   "Ӧ�ó�������"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   6
      Tag             =   "Ӧ�ó�������"
      Top             =   1125
      Width           =   4092
   End
   Begin VB.Label lblTitle 
      Caption         =   "Ӧ�ó������"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Tag             =   "Ӧ�ó������"
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
      Caption         =   "�汾"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Tag             =   "�汾"
      Top             =   780
      Width           =   4092
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "���棺��������Apache Version 2.0��ԴЭ�鱣������ϸ���ݲμ�http://www.apache.org/licenses/LICENSE-2.0"
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   3
      Tag             =   "����: ..."
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ע�����ȫѡ��...
Const KEY_ALL_ACCESS = &H2003F
                                          

' ע���������...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode �ս�β�ַ���
Const REG_DWORD = 4                      ' 32λ��


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
    lblVersion.Caption = "�汾 " & App.Major & "." & App.Minor & "." & App.Revision
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
        

        ' ��ע�����ϵͳ��Ϣ����·��\����...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' ����ע�����ϵͳ��Ϣ����·��...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' ��֤��֪�� 32 λ�ļ��汾�Ĵ���
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' ���� - �ļ��Ҳ���...
                Else
                        GoTo SysInfoErr
                End If
        ' ���� - ע������Ҳ���...
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
        Dim hKey As Long                                        ' �򿪵�ע�������
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' ע������������
        Dim tmpVal As String                                    ' ��ʱ�洢һ��ע����ֵ
        Dim KeyValSize As Long                                  ' ע����������С
        '------------------------------------------------------------
        ' �ڼ���{HKEY_LOCAL_MACHINE...}֮�´�ע���
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע����
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������...
        

        tmpVal = String$(1024, 0)                             ' ��������ռ�
        KeyValSize = 1024                                       ' ��Ǳ�����С
        

        '------------------------------------------------------------
        ' ����ע����ֵ...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' ���/������ֵ
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������
      

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)


        '------------------------------------------------------------
        ' ����ת���ļ�ֵ����...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' ������������...
        Case REG_SZ                                             ' �ַ���ע������������
                KeyVal = tmpVal                                     ' �����ַ���ֵ
        Case REG_DWORD                                          ' ˫����ע������������
                For i = Len(tmpVal) To 1 Step -1                    ' ת��ÿһҳ
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' һ���ַ�һ���ַ�������ֵ
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' ת��˫����Ϊ�ַ���
        End Select
        

        GetKeyValue = True                                      ' ���سɹ�
        rc = RegCloseKey(hKey)                                  ' �ر�ע����
        Exit Function                                           ' �˳�
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' �践��ֵΪ���ַ���
        GetKeyValue = False                                     ' ����ʧ��
        rc = RegCloseKey(hKey)                                  ' �ر�ע����
End Function

