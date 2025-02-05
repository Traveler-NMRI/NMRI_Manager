VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登录"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Tag             =   "登录"
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   360
      Left            =   2100
      TabIndex        =   5
      Tag             =   "取消"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   360
      Left            =   495
      TabIndex        =   4
      Tag             =   "确定"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1305
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   270
      Left            =   1305
      TabIndex        =   3
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "密码(&P):"
      Height          =   248
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Tag             =   "密码(&P):"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "用户名(&U):"
      Height          =   248
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Tag             =   "用户名(&U):"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long


Public OK As Boolean
Private Sub Form_Load()
    Dim sBuffer As String
    Dim lSize As Long


    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        txtUserName.Text = Left$(sBuffer, lSize)
    Else
        txtUserName.Text = vbNullString
    End If
End Sub



Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    'ToDo: 创建测试密码是否正确
    '检查正确密码
    If txtPassword.Text = "" Then
        OK = True
        Me.Hide
    Else
        MsgBox "密码错误，再试一次！", , "登录"
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    End If
End Sub

