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
            Key             =   "�½�"
            Object.ToolTipText     =   "�µ�ɨ��"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "��"
            Object.ToolTipText     =   "��"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ɾ��"
            Object.ToolTipText     =   "ɾ��"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Drawing"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "���"
            Object.ToolTipText     =   "���"
            ImageKey        =   "Sum"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "�����"
            Object.ToolTipText     =   "�����"
            ImageKey        =   "Camera"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
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
            Text            =   "״̬"
            TextSave        =   "״̬"
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
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuScanNew 
         Caption         =   "�µ�ɨ��(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "��(&O)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "�ر�(&C)"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "���Ϊ(&A)..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "ȫ������(&L)"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "����(&I)"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "����(&D)..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "��ͼ(&V)"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "������(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "״̬��(&B)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "����(&T)"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "ѡ��(&O)..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "����(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "���(&C)"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "����ƽ��(&H)"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "����ƽ��(&V)"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "����ͼ��(&A)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Ŀ¼(&C)"
      End
      Begin VB.Menu mnuHelpTipDaily 
         Caption         =   "�ջ�����(&P)"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "������������(&S)..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A) "
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
        Case "�½�"
            mnuScanNew_Click
        Case "��"
            MsgBox "��� '��' ��ť���롣"
        Case "����"
            MsgBox "��� '����' ��ť���롣"
        Case "ɾ��"
            MsgBox "��� 'ɾ��' ��ť���롣"
        Case "����"
            MsgBox "��� '����' ��ť���롣"
        Case "���"
            MsgBox "��� '���' ��ť���롣"
        Case "�����"
            MsgBox "��� '�����' ��ť���롣"
        Case "����"
            mnuFileProperties_Click
    End Select
End Sub
Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub
Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer
    '����������û�а����ļ�����ʾ��Ϣ���û�
    '�����ڡ��������ԡ��Ի�����ΪӦ�ó������ð����ļ�
    If Len(App.HelpFile) = 0 Then
        MsgBox "�޷���ʾ����Ŀ¼���ù���û��������İ�����", vbInformation, Me.Caption
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
    '����������û�а����ļ�����ʾ��Ϣ���û�
    '�����ڡ��������ԡ��Ի�����ΪӦ�ó������ð����ļ�
    If Len(App.HelpFile) = 0 Then
        MsgBox "�޷���ʾ����Ŀ¼���ù���û��������İ�����", vbInformation, Me.Caption
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
    MsgBox "��� 'mnuFileSend_Click' ���롣"
End Sub

Private Sub mnuFileProperties_Click()
    MsgBox "��� 'mnuFileProperties_Click' ���롣"
End Sub

Private Sub mnuFileSaveAll_Click()
    MsgBox "��� 'mnuFileSaveAll_Click' ���롣"
End Sub

Private Sub mnuFileSaveAs_Click()
    MsgBox "��� 'mnuFileSaveAs_Click' ���롣"
End Sub

Private Sub mnuFileSave_Click()
    MsgBox "��� 'mnuFileSave_Click' ���롣"
End Sub

Private Sub mnuFileClose_Click()
    MsgBox "��� 'mnuFileClose_Click' ���롣"
End Sub

Private Sub mnuFileOpen_Click()
    MsgBox "��� 'mnuFileOpen_Click' ���롣"
End Sub
