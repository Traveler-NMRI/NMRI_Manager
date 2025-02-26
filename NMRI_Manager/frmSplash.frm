VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4590
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   7380
      Begin VB.PictureBox picLogo 
         Height          =   2385
         Left            =   510
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   2325
         ScaleWidth      =   1755
         TabIndex        =   2
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "使用权属于Traveler-NMRI"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   1
         Tag             =   "使用权属于"
         Top             =   300
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "产品"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   32.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2670
         TabIndex        =   8
         Tag             =   "产品"
         Top             =   1200
         Width           =   1320
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Made by Traveler-NMRI"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2505
         TabIndex        =   7
         Tag             =   "公司产品"
         Top             =   765
         Width           =   4095
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Windows NT"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5205
         TabIndex        =   6
         Tag             =   "平台"
         Top             =   2400
         Width           =   1800
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "版本"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6495
         TabIndex        =   5
         Tag             =   "版本"
         Top             =   2760
         Width           =   510
      End
      Begin VB.Label lblWarning 
         Caption         =   "警告：本程序受Apache Version 2.0开源协议保护，详细内容参见http://www.apache.org/licenses/LICENSE-2.0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   3
         Tag             =   "警告"
         Top             =   3720
         Width           =   6855
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright [2024] Traveler-NMRI"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4350
         TabIndex        =   4
         Tag             =   "版权"
         Top             =   3120
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

