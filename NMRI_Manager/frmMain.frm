VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "NMRI_Manager"
   ClientHeight    =   2505
   ClientLeft      =   7665
   ClientTop       =   3945
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "新建"
            Object.ToolTipText     =   "新的扫描"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "打开"
            Object.ToolTipText     =   "打开"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "保存"
            Object.ToolTipText     =   "保存"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "删除"
            Object.ToolTipText     =   "删除"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "绘制"
            Object.ToolTipText     =   "绘制"
            ImageKey        =   "Drawing"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "求和"
            Object.ToolTipText     =   "求和"
            ImageKey        =   "Sum"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "照相机"
            Object.ToolTipText     =   "照相机"
            ImageKey        =   "Camera"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "属性"
            Object.ToolTipText     =   "属性"
            ImageKey        =   "Properties"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2235
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2593
            Text            =   "状态"
            TextSave        =   "状态"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2024-8-15"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "21:26"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2400
      Top             =   1305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0336
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0448
            Key             =   "Drawing"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":055A
            Key             =   "Sum"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066C
            Key             =   "Camera"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":077E
            Key             =   "Properties"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuScanNew 
         Caption         =   "新的扫描(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "打开(&O)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "关闭(&C)"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "另存为(&A)..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "全部保存(&L)"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "属性(&I)"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "发送(&D)..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "视图(&V)"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "工具栏(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "状态栏(&B)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "选项(&O)..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "窗口(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "层叠(&C)"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "横向平铺(&H)"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "纵向平铺(&V)"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "排列图标(&A)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "目录(&C)"
      End
      Begin VB.Menu mnuHelpTipDaily 
         Caption         =   "日积月累(&P)"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "搜索帮助主题(&S)..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A) "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
End Sub


Private Sub LoadNewScan()
    Dim frmD As frmDocument
    Set frmD = New frmDocument
    frmD.Caption = "NMR Scan Workspace(Scan Name:"
    frmD.Show
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuHelpTipDaily_Click()
frmTip.Show vbModal, Me
End Sub

Private Sub mnuScanNew_Click()
frmNew.Show vbModal, Me
LoadNewScan
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "新建"
            mnuScanNew_Click
        Case "打开"
            MsgBox "添加 '打开' 按钮代码。"
        Case "保存"
            MsgBox "添加 '保存' 按钮代码。"
        Case "删除"
            MsgBox "添加 '删除' 按钮代码。"
        Case "绘制"
            MsgBox "添加 '绘制' 按钮代码。"
        Case "求和"
            MsgBox "添加 '求和' 按钮代码。"
        Case "照相机"
            MsgBox "添加 '照相机' 按钮代码。"
        Case "属性"
            mnuFileProperties_Click
    End Select
End Sub
Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub
Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer
    '如果这个工程没有帮助文件，显示消息给用户
    '可以在“工程属性”对话框中为应用程序设置帮助文件
    If Len(App.HelpFile) = 0 Then
        MsgBox "无法显示帮助目录，该工程没有相关联的帮助。", vbInformation, Me.Caption
    Else
    On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    '如果这个工程没有帮助文件，显示消息给用户
    '可以在“工程属性”对话框中为应用程序设置帮助文件
    If Len(App.HelpFile) = 0 Then
        MsgBox "无法显示帮助目录，该工程没有相关联的帮助。", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub
Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub
Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub
Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub
Private Sub mnuToolsOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSend_Click()
    MsgBox "添加 'mnuFileSend_Click' 代码。"
End Sub

Private Sub mnuFileProperties_Click()
    MsgBox "添加 'mnuFileProperties_Click' 代码。"
End Sub

Private Sub mnuFileSaveAll_Click()
    MsgBox "添加 'mnuFileSaveAll_Click' 代码。"
End Sub

Private Sub mnuFileSaveAs_Click()
    MsgBox "添加 'mnuFileSaveAs_Click' 代码。"
End Sub

Private Sub mnuFileSave_Click()
    MsgBox "添加 'mnuFileSave_Click' 代码。"
End Sub

Private Sub mnuFileClose_Click()
    MsgBox "添加 'mnuFileClose_Click' 代码。"
End Sub

Private Sub mnuFileOpen_Click()
    MsgBox "添加 'mnuFileOpen_Click' 代码。"
End Sub
