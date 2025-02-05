VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNew 
   Caption         =   "新建扫描向导"
   ClientHeight    =   4680
   ClientLeft      =   6810
   ClientTop       =   3030
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   6525
   Begin VB.CommandButton DownCommand 
      Height          =   255
      Left            =   4320
      Picture         =   "frmNew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton UpCommand 
      Height          =   255
      Left            =   4080
      Picture         =   "frmNew.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton DeleteCommand 
      Height          =   255
      Left            =   3720
      Picture         =   "frmNew.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton AddCommand 
      Height          =   255
      Left            =   3480
      Picture         =   "frmNew.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   3480
      TabIndex        =   26
      Top             =   1370
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "顺序"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "序列"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   140
      ImageHeight     =   260
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNew.frx":0528
            Key             =   "Wizard"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNew.frx":4EDA
            Key             =   "Finish"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "高级扫描设置"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   3690
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置"
      Height          =   255
      Left            =   5640
      TabIndex        =   23
      Top             =   1050
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   3480
      TabIndex        =   22
      Top             =   680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Height          =   1215
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   20
      Top             =   2850
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   3480
      TabIndex        =   18
      Top             =   2490
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   3480
      TabIndex        =   16
      Top             =   2130
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   3480
      TabIndex        =   14
      Top             =   1770
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frmNew.frx":988C
      Left            =   3480
      List            =   "frmNew.frx":98A8
      TabIndex        =   12
      Text            =   "1-H Spectroscopy"
      Top             =   1370
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "浏览"
      Height          =   255
      Left            =   5640
      TabIndex        =   10
      Top             =   1050
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3480
      TabIndex        =   9
      Text            =   "C:\NMRI_Manager\Scan_Name\"
      Top             =   1050
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   3480
      TabIndex        =   7
      Top             =   690
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton CancleCommand 
      Caption         =   "取消"
      Height          =   312
      Left            =   1560
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Tag             =   "104"
      Top             =   4320
      Width           =   1092
   End
   Begin VB.CommandButton BackCommand 
      Caption         =   "上一步(&B)"
      Enabled         =   0   'False
      Height          =   312
      Left            =   2760
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Tag             =   "104"
      Top             =   4320
      Width           =   1092
   End
   Begin VB.CommandButton NextCommand 
      Caption         =   "下一步(&N)"
      Height          =   312
      Left            =   3840
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Tag             =   "104"
      Top             =   4320
      Width           =   1092
   End
   Begin VB.CommandButton FinishCommand 
      Caption         =   "完成(&F)"
      Enabled         =   0   'False
      Height          =   312
      Left            =   5280
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Tag             =   "104"
      Top             =   4320
      Width           =   1092
   End
   Begin VB.Label Label11 
      Caption         =   "扫描序列:"
      Height          =   255
      Left            =   2520
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "设备:"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "备注:"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "检测人员:"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "样本时间:"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "样品名称:"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "扫描类型:"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "保存路径:"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "扫描名称:"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "NMRI_Manager"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6360
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label1 
      Caption         =   "此向导将帮助您创建一个新的扫描"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   3900
      Left            =   120
      Picture         =   "frmNew.frx":992E
      Top             =   120
      Width           =   2100
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim step As Integer
Dim isSpec As Boolean
Private Sub AddCommand_Click()
frmAddSeq.isSpecOrImage isSpec
frmAddSeq.Show vbModal, Me
End Sub

Private Sub BackCommand_Click()
step = step - 1
SetWizard
End Sub

Private Sub CancleCommand_Click()
Unload Me
End Sub

Private Sub FinishCommand_Click()
Unload Me
End Sub

Private Sub Form_Load()
step = 0
Combo1.ListIndex = 0
End Sub


Private Sub NextCommand_Click()
step = step + 1
SetWizard
End Sub
Private Function SetWizard()
If step = 0 Then
Label1.Caption = "此向导将帮助您创建一个新的扫描"
Image1.Picture = ImageList1.ListImages.Item(1).Picture
ListView1.Visible = False
NextCommand.Enabled = True
BackCommand.Enabled = False
FinishCommand.Enabled = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Combo1.Visible = False
Combo2.Visible = False
Command1.Visible = False
Command2.Visible = False
Command4.Visible = False
AddCommand.Visible = False
DeleteCommand.Visible = False
UpCommand.Visible = False
DownCommand.Visible = False
ElseIf step = 1 Then
Label1.Caption = "提供扫描名称、文件保存路径、类型、样品等信息"
ListView1.Visible = False
Image1.Picture = ImageList1.ListImages.Item(1).Picture
NextCommand.Enabled = True
BackCommand.Enabled = True
FinishCommand.Enabled = False
AddCommand.Visible = False
DeleteCommand.Visible = False
UpCommand.Visible = False
DownCommand.Visible = False
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = False
Label11.Visible = False
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Combo1.Visible = True
Combo2.Visible = False
Command1.Visible = True
Command2.Visible = False
Command4.Visible = False
ElseIf step = 2 Then
Label1.Caption = "提供设备、扫描序列等信息"
ListView1.Visible = True
Image1.Picture = ImageList1.ListImages.Item(1).Picture
NextCommand.Enabled = True
BackCommand.Enabled = True
FinishCommand.Enabled = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = True
Label11.Visible = True
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Combo1.Visible = False
Combo2.Visible = True
Command1.Visible = False
Command2.Visible = True
Command4.Visible = True
AddCommand.Visible = True
DeleteCommand.Visible = True
UpCommand.Visible = True
DownCommand.Visible = True
If Combo1.ListIndex < 4 And Combo1.ListIndex >= 0 Then
isSpec = True
Else
isSpec = False
End If
ElseIf step = 3 Then
Label1.Caption = "新的扫描已创建！完成创建以进行扫描！"
ListView1.Visible = False
Image1.Picture = ImageList1.ListImages.Item(2).Picture
FinishCommand.Enabled = True
NextCommand.Enabled = False
BackCommand.Enabled = True
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Combo1.Visible = False
Combo2.Visible = False
Command1.Visible = False
Command2.Visible = False
Command4.Visible = False
AddCommand.Visible = False
DeleteCommand.Visible = False
UpCommand.Visible = False
DownCommand.Visible = False
End If
End Function
